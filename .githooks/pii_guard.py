#!/usr/bin/env python3
"""ด่านกันข้อมูลส่วนบุคคลหลุดขึ้น GitHub (repo นี้เป็นสาธารณะ ใครก็เปิดดูได้)

รันเองอัตโนมัติก่อน commit (pre-commit) และก่อน push (pre-push) — เจอของต้องห้ามจะบล็อกทันที
  1) ไฟล์ต้องอยู่ในรายการที่อนุญาต (pii_allowlist.json → paths) เท่านั้น ไฟล์คดี/ตัวอย่างจริงเข้าไม่ได้
  2) เลขบัตรประชาชน (ตรวจ checksum จริง) / เบอร์มือถือ / เลขบัญชีธนาคาร — ทั้งในโค้ด ในไฟล์ Word
     และในแม่แบบที่ฝังเป็น base64 อยู่ใน HTML
  3) รูปทุกรูป รวมรูปที่ซ่อนอยู่ใน .docx (เช่นลายเซ็นที่ค้างในแพ็กเกจ) ต้องตรงกับรายการรูปที่อนุญาต
  4) ชื่อคนจริงในแม่แบบ Word, ชื่อผู้สร้าง/ผู้แก้ไขในข้อมูลแฝงของไฟล์

ใช้งาน:
  python .githooks/pii_guard.py staged         ตรวจไฟล์ที่ stage ไว้ (pre-commit เรียกให้)
  python .githooks/pii_guard.py push           ตรวจทุก commit ที่กำลังจะ push (pre-push เรียกให้)
  python .githooks/pii_guard.py tree [rev]     ตรวจไฟล์ทั้งหมดใน commit นั้น (ค่าเริ่มต้น HEAD)
  python .githooks/pii_guard.py history        ตรวจทุกไฟล์ทุกเวอร์ชันในประวัติทั้งหมด
  python .githooks/pii_guard.py hash <ไฟล์>    พิมพ์ sha256 ของรูป เพื่อเพิ่มเข้า media_sha256 (ต้องดูรูปก่อนทุกครั้ง)
"""
import base64
import fnmatch
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import zipfile

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ALLOW = json.load(io.open(os.path.join(HERE, "pii_allowlist.json"), encoding="utf-8"))
ZEROS = "0" * 40

IMAGE_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".webp", ".svg", ".heic")
ZIP_EXT = (".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".odt", ".ods", ".zip")
THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")


def git(*args, data=False):
    out = subprocess.run(["git", *args], capture_output=True, check=True).stdout
    return out if data else out.decode("utf-8", "replace")


# ---------------------------------------------------------------- ตรวจข้อความ
ID_RE = re.compile(r"(?<![\d])(\d)[ \-]?(\d{4})[ \-]?(\d{5})[ \-]?(\d{2})[ \-]?(\d)(?![\d])")
MOBILE_RE = re.compile(r"(?<![\d])0[689]\d[ \-]?\d{3}[ \-]?\d{4}(?![\d])")
BANK_RE = re.compile(r"(?:บัญชี|account|acct)[^\d\n<>{}]{0,30}?((?:\d[ \-]?){10,12})(?![\d])", re.I)
TITLE_NAME_RE = re.compile(r"(นางสาว|นาง|นาย|น\.ส\.|ด\.ช\.|ด\.ญ\.)\s*([ก-ฮ][ก-๙]+)(\s+[ก-ฮ][ก-๙]+)?")
PAREN_NAME_RE = re.compile(r"\(\s*([ก-ฮ][ก-๙]{1,})\s+([ก-ฮ][ก-๙]{1,})\s*\)")
# คำที่ขึ้นต้นด้วย "นาย/นาง" แต่ไม่ใช่ชื่อคน
NOT_A_NAME = ("นายายอาม", "นายอำเภอ", "นายก", "นายแพทย์", "นายตำรวจ", "นายทหาร", "นายจ้าง", "นายหน้า",
              "นายประกัน", "นายทะเบียน", "นายด่าน", "นายสิบ", "นายร้อย", "นายพล", "นายช่าง", "นายงาน",
              "นายเวร", "นายท่า", "นายอากร", "นางพยาบาล", "นางงาม")


def thai_id_valid(d):
    return len(d) == 13 and d[0] != "0" and (11 - sum(int(d[i]) * (13 - i) for i in range(12)) % 11) % 10 == int(d[12])


def mask(s):
    s = s.strip()
    return s[:4] + "•" * max(0, len(s) - 6) + s[-2:] if len(s) > 6 else "•" * len(s)


def allowed_text(snippet):
    return any(a in snippet for a in ALLOW.get("allowed_text", []))


def scan_text(text, where, template=False):
    found = []
    t = text.translate(THAI_DIGITS)
    for m in ID_RE.finditer(t):
        digits = "".join(m.groups())
        if thai_id_valid(digits) and not allowed_text(m.group(0)):
            found.append(f"{where}: เลขบัตรประชาชน {mask(m.group(0))}")
    for m in MOBILE_RE.finditer(t):
        if not allowed_text(m.group(0)):
            found.append(f"{where}: เบอร์มือถือ {mask(m.group(0))}")
    for m in BANK_RE.finditer(t):
        if not allowed_text(m.group(1)):
            found.append(f"{where}: เลขบัญชี {mask(m.group(1))}")
    if template:
        for m in TITLE_NAME_RE.finditer(t):
            word = m.group(1) + m.group(2)
            if word.startswith(NOT_A_NAME) or allowed_text(m.group(0)):
                continue
            found.append(f"{where}: ชื่อคน \"{m.group(1)}{m.group(2)[:2]}…\" ในแม่แบบ (แม่แบบต้องมีแต่แท็ก {{...}})")
        for m in PAREN_NAME_RE.finditer(t):
            if not allowed_text(m.group(0)):
                found.append(f"{where}: ชื่อ-สกุลในวงเล็บ \"( {m.group(1)[:2]}… )\" ในแม่แบบ")
    return found


# ---------------------------------------------------------------- ตรวจรูป / ไฟล์ Word
def sniff_image(b):
    return (b[:8] == b"\x89PNG\r\n\x1a\n" or b[:3] == b"\xff\xd8\xff" or b[:4] in (b"GIF8", b"II*\x00", b"MM\x00*")
            or b[:2] == b"BM" or b[:4] in (b"\xd7\xcd\xc6\x9a", b"\x01\x00\x00\x00") or b[8:12] == b"WEBP")


def check_media(data, where):
    h = hashlib.sha256(data).hexdigest()
    if h in ALLOW.get("media_sha256", {}):
        return []
    return [f"{where}: รูป/ไฟล์ฝังที่ไม่อยู่ในรายการอนุญาต (sha256 {h[:12]}…) — อาจเป็นลายเซ็น/รูปถ่ายคนจริง"]


def xml_text(xml):
    xml = re.sub(r"</w:p>|</a:p>|<w:br/>|<w:tab/>|</si>|</c>|</row>", "\n", xml)
    return re.sub(r"<[^>]+>", "", xml)


def scan_zip(data, where, template=False):
    found = []
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return [f"{where}: ไฟล์ zip/docx เสีย ตรวจไม่ได้ — ไม่อนุญาต"]
    for name in z.namelist():
        part = z.read(name)
        w = f"{where} › {name}"
        low = name.lower()
        if low.endswith(IMAGE_EXT) or low.endswith(".bin") or "thumbnail" in low or sniff_image(part[:16]):
            found += check_media(part, w)
            continue
        if not (low.endswith(".xml") or low.endswith(".rels")):
            found += check_media(part, w)
            continue
        xml = part.decode("utf-8", "replace")
        if low.startswith("docprops/"):
            for tag in ("dc:creator", "cp:lastModifiedBy", "Company", "Manager", "dc:title", "dc:subject", "cp:keywords", "dc:description"):
                for val in re.findall(rf"<{tag}[^>]*>([^<]*)</{tag}>", xml):
                    if val.strip() and val.strip() not in ALLOW.get("allowed_meta", []):
                        found.append(f"{w}: ข้อมูลแฝง {tag} = \"{val.strip()[:3]}…\" (ชื่อผู้สร้าง/ผู้แก้ไขไฟล์)")
        for val in set(re.findall(r'w:author="([^"]*)"', xml)):
            if val.strip() and val.strip() not in ALLOW.get("allowed_meta", []):
                found.append(f"{w}: ชื่อผู้แก้ไข (w:author) \"{val[:3]}…\" ค้างในไฟล์")
        found += scan_text(xml_text(xml), w, template=template and low.startswith("word/"))
    return found


B64_RUN = re.compile(r"[A-Za-z0-9+/]{120,}={0,2}")


def scan_blob(path, data):
    low = path.lower()
    if low.endswith(ZIP_EXT) or data[:4] == b"PK\x03\x04":
        return scan_zip(data, path, template=fnmatch.fnmatch(path, "build/*_template.docx"))
    if low.endswith(IMAGE_EXT) or sniff_image(data[:16]):
        return check_media(data, path)
    text = data.decode("utf-8", "replace")
    found = []
    # ก้อน base64 ใน HTML/JS: แม่แบบ Word ที่ฝังไว้ หรือรูป — ถอดแล้วตรวจข้างในด้วย
    for m in B64_RUN.finditer(text):
        blob = m.group(0)
        try:
            raw = base64.b64decode(blob + "=" * (-len(blob) % 4))
        except Exception:
            continue
        key = re.search(r"(\w+)\s*:\s*[\"']?$", text[max(0, m.start() - 60):m.start()])
        w = f"{path} › base64 {key.group(1) if key else '@' + str(m.start())}"
        if raw[:4] == b"PK\x03\x04":
            found += scan_zip(raw, w, template=True)
        elif sniff_image(raw[:16]):
            found += check_media(raw, w)
        elif re.match(rb"\s*(<\?xml|<svg)", raw[:64]):
            # รูป SVG (เช่นลายน้ำ) เป็นตัวหนังสือ — อาจมีชื่อคนซ่อนอยู่ ต้องผ่านการตรวจด้วยตาเหมือนรูปอื่น
            found += check_media(raw, w)
            found += scan_text(raw.decode("utf-8", "replace"), w)
    # SVG ที่ฝังเป็น data URI แบบไม่เข้ารหัส base64 (ตัว svg ต่อท้ายเครื่องหมายจุลภาคตรง ๆ) ก็ต้องผ่านการอนุญาตเหมือนกัน
    for m in re.finditer(r"data:image/svg\+xml(?:;charset=[\w-]+)?,([^\"')]+)", text):
        found += check_media(m.group(1).encode("utf-8"), f"{path} › svg@{m.start()}")
    found += scan_text(B64_RUN.sub(" ", text), path)
    return found


def path_allowed(path):
    return any(fnmatch.fnmatch(path, p) for p in ALLOW["paths"]) and \
        not any(fnmatch.fnmatch(path, p) for p in ALLOW.get("paths_denied", []))


def scan_files(items):
    """items: [(path, bytes)]"""
    found = []
    for path, data in items:
        if not path_allowed(path):
            found.append(f"{path}: ไฟล์นี้ไม่อยู่ในรายการไฟล์ที่อนุญาตให้ขึ้น GitHub")
            continue
        found += scan_blob(path, data)
    return found


# ---------------------------------------------------------------- โหมดต่าง ๆ
def staged_items():
    names = [n for n in git("diff", "--cached", "--name-only", "--diff-filter=ACMRT", "-z").split("\0") if n]
    return [(n, git("show", f":{n}", data=True)) for n in names]


def commit_items(rev, changed_only=True):
    if changed_only:
        names = git("diff-tree", "-r", "--root", "-m", "--no-commit-id", "--name-only", "--diff-filter=ACMRT", "-z", rev).split("\0")
    else:
        names = git("ls-tree", "-r", "--name-only", "-z", rev).split("\0")
    return [(n, git("show", f"{rev}:{n}", data=True)) for n in dict.fromkeys(n for n in names if n)]


def report(found, title):
    if not found:
        print(f"✓ pii_guard: {title} — ไม่พบข้อมูลส่วนบุคคล")
        return 0
    print(f"✗ pii_guard: {title} — บล็อกไว้ พบ {len(found)} รายการ:", file=sys.stderr)
    for f in found[:60]:
        print("   - " + f, file=sys.stderr)
    if len(found) > 60:
        print(f"   ... และอีก {len(found) - 60} รายการ", file=sys.stderr)
    print("   ลบข้อมูลจริงออกก่อน — ห้ามใช้ --no-verify ข้ามด่านนี้", file=sys.stderr)
    return 1


def main(argv):
    os.chdir(git("rev-parse", "--show-toplevel").strip())
    mode = argv[1] if len(argv) > 1 else "staged"
    if mode == "staged":
        return report(scan_files(staged_items()), "ไฟล์ที่จะ commit")
    if mode == "tree":
        rev = argv[2] if len(argv) > 2 else "HEAD"
        return report(scan_files(commit_items(rev, changed_only=False)), f"ไฟล์ทั้งหมดใน {rev}")
    if mode == "push":
        found = []
        for line in sys.stdin.read().splitlines():
            parts = line.split()
            if len(parts) < 4 or parts[1] == ZEROS:
                continue
            for c in git("rev-list", parts[1], "--not", "--remotes").split():
                found += [f"[{c[:7]}] {f}" for f in scan_files(commit_items(c))]
            found += [f"[{parts[1][:7]} ทั้งต้นไม้] {f}" for f in scan_files(commit_items(parts[1], changed_only=False))]
        return report(sorted(set(found)), "ทุก commit ที่จะ push")
    if mode == "history":
        found = []
        for c in git("rev-list", "--all").split():
            found += [f"[{c[:7]}] {f}" for f in scan_files(commit_items(c))]
        return report(sorted(set(found)), "ประวัติทั้งหมด")
    if mode == "hash":
        for p in argv[2:]:
            print(hashlib.sha256(open(p, "rb").read()).hexdigest(), p)
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
