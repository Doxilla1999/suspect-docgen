"""
Combines the generated .docx templates (in build/) with app_template.html
into a single self-contained suspect-docgen.html file.

Run generate_templates.py first, then this script:
    python3 generate_templates.py
    python3 assemble.py
"""
import base64
import glob
import os
import re
import sys
import zipfile
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BUILD_DIR = os.path.join(SCRIPT_DIR, "build")

# ข้อมูลแฝงในไฟล์ Word ที่อาจมีชื่อคน/ชื่อเครื่องของผู้ทำฟอร์มต้นฉบับ — ล้างทิ้งก่อนฝังลงหน้าเว็บสาธารณะ
CORE_FIELDS = ("dc:creator", "cp:lastModifiedBy", "dc:title", "dc:subject", "dc:description",
               "cp:keywords", "cp:category", "cp:contentStatus")
APP_FIELDS = ("Company", "Manager")


def scrub_docx(path):
    """ล้างชื่อผู้สร้าง/ผู้แก้ไข/บริษัท และลบรูปย่อหน้าแรก (thumbnail) ออกจากแม่แบบ — ไม่แตะเนื้อเอกสาร
    เขียนไฟล์ใหม่เฉพาะเมื่อมีอะไรต้องล้างจริง (รันซ้ำแล้วไฟล์ไม่เปลี่ยน)"""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    out, changed = [], False
    for info, data in items:
        name = info.filename
        if name.lower().startswith("docprops/thumbnail"):
            changed = True
            continue
        if name in ("docProps/core.xml", "docProps/app.xml", "_rels/.rels", "[Content_Types].xml"):
            xml = data.decode("utf-8")
            new = xml
            for tag in (CORE_FIELDS if name == "docProps/core.xml" else APP_FIELDS if name == "docProps/app.xml" else ()):
                new = re.sub(rf"(<{tag}\b[^>]*>)[^<]*(</{tag}>)", r"\1\2", new)
            new = re.sub(r'<(?:Relationship|Override)\b[^>]*"/?docProps/thumbnail[^"]*"[^>]*/>', "", new)
            if new != xml:
                changed = True
                data = new.encode("utf-8")
        out.append((info, data))
    if changed:
        tmp = path + ".tmp"
        with zipfile.ZipFile(tmp, "w") as z:
            for info, data in out:
                z.writestr(info, data)
        os.replace(tmp, path)
    return changed

TEMPLATE_FILES = {
    "__B64_DRUG115__": "drug_test_115_template.docx",
    "__B64_URINE__": "urine_referral_template.docx",
    "__B64_COURT__": "court_referral_template.docx",
    "__B64_M80__": "phone_consent_m80_template.docx",
    "__B64_PROFILE_PHOTOS__": "profile_and_photos_template.docx",
    "__B64_ONLINE_REPORT__": "online_report_template.docx",
}


def b64_of(filename):
    path = os.path.join(BUILD_DIR, filename)
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def main():
    for path in sorted(glob.glob(os.path.join(BUILD_DIR, "*_template.docx"))):
        if scrub_docx(path):
            print(f"scrubbed metadata: {os.path.basename(path)}")

    html_path = os.path.join(SCRIPT_DIR, "app_template.html")
    with open(html_path, encoding="utf-8") as f:
        html = f.read()

    for token, filename in TEMPLATE_FILES.items():
        html = html.replace(token, b64_of(filename))

    # ตราเวลา build (พ.ศ.) โชว์ที่หัวหน้าเว็บ ไว้เช็คว่าเบราว์เซอร์โหลดเวอร์ชันล่าสุดแล้วหรือยัง
    now = datetime.now()
    html = html.replace("__BUILD_STAMP__", f"{now.day:02d}/{now.month:02d}/{now.year + 543} {now:%H:%M}")

    if "__B64_" in html:
        raise SystemExit("Leftover __B64_ token found — a template failed to embed. Aborting.")

    out_path = os.path.join(SCRIPT_DIR, "suspect-docgen.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"assembled {out_path} ({len(html):,} chars)")

    # ตรวจข้อมูลส่วนบุคคลทันที (ด่านเดียวกับ pre-commit) จะได้รู้ตั้งแต่ตอน build ไม่ต้องรอ commit
    sys.path.insert(0, os.path.join(SCRIPT_DIR, ".githooks"))
    import pii_guard
    files = ["suspect-docgen.html", "app_template.html"] + \
        [os.path.relpath(p, SCRIPT_DIR).replace(os.sep, "/") for p in sorted(glob.glob(os.path.join(BUILD_DIR, "*_template.docx")))
         if not os.path.basename(p).startswith("_superseded_")]  # ไฟล์เก่าที่ gitignore ไว้ ไม่ขึ้น GitHub
    found = pii_guard.scan_files([(f, open(os.path.join(SCRIPT_DIR, f), "rb").read()) for f in files])
    pii_guard.report(found, "ตรวจข้อมูลส่วนบุคคลหลัง build")
    if found:
        print("!!! อย่า commit จนกว่าจะแก้รายการข้างบน (pre-commit จะบล็อกอยู่แล้ว)", file=sys.stderr)


if __name__ == "__main__":
    main()
