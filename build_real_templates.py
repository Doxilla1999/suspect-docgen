"""แปลงไฟล์ฟอร์มจริงของสถานี (forms/*.docx) ให้เป็นแม่แบบที่มีแท็ก {...} โดยไม่แตะเลย์เอาต์เดิม

ต่างจาก generate_templates.py ที่ "สร้างเอกสารขึ้นมาใหม่" ทั้งฉบับ — สคริปต์นี้เปิดไฟล์ฟอร์มจริง
แล้วแทนที่เฉพาะ "ข้อความในช่องที่ต้องกรอก" ด้วยแท็ก ทำให้ฟอนต์/ระยะ/ระยะบรรทัด/ตาราง
ของฟอร์มเดิมคงอยู่ครบ เวลา export ออกมาฟอร์มจึงไม่ขยับ
"""
import copy
import os
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

# ตำแหน่งบล็อกหัวจดหมายด้านขวา (ชื่อสถานี / จังหวัด) วัดจากขอบซ้ายของพื้นที่พิมพ์
LETTERHEAD_TAB = Cm(10)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FORMS = os.path.join(SCRIPT_DIR, "forms") + os.sep
OUT = os.path.join(SCRIPT_DIR, "build") + os.sep


def set_runs(para, spans):
    """เขียนข้อความลง run ตามดัชนีที่ระบุ แล้วล้าง run ที่เหลือในกลุ่มเดียวกัน

    spans: dict {run_index: text} — run ที่ไม่ได้ระบุจะถูกปล่อยไว้เหมือนเดิม
    ใช้วิธีนี้เพื่อคงรูปแบบ (ตัวหนา/ฟอนต์/แท็บ) ของ run ที่ไม่เกี่ยวข้องไว้
    """
    for idx, text in spans.items():
        para.runs[idx].text = text


def clear_runs(para, indexes):
    for idx in indexes:
        para.runs[idx].text = ""


def flatten(para, text):
    """ยุบทุก run ในย่อหน้าให้เหลือ run แรกอันเดียวแล้วใส่ข้อความใหม่

    ใช้กับย่อหน้าที่ทุก run มีรูปแบบเดียวกันอยู่แล้ว ข้อดีคือแท็กจะไม่ถูก Word
    ตัดคร่อมหลาย run ซึ่งทำให้แทนที่ข้อความตอนสร้างเอกสารไม่สำเร็จ
    """
    para.runs[0].text = text
    for r in para.runs[1:]:
        r.text = ""


def replace_across_runs(para, old, new):
    """แทนที่ข้อความในย่อหน้า แม้ Word จะตัดข้อความนั้นคร่อมหลาย run

    คงรูปแบบของ run แรกที่ทับกับข้อความเดิมไว้ (ตัวหนา/ฟอนต์) แล้วล้างส่วนที่ทับใน run ถัดๆ ไป
    ทำให้แท็กที่ใส่เข้าไปอยู่ใน run เดียว ไม่ถูกตัดคร่อมตอนแทนค่าในเว็บแอป
    """
    runs = para.runs
    full = "".join(r.text for r in runs)
    idx = full.find(old)
    if idx < 0:
        raise ValueError(f"ไม่พบข้อความ {old!r} ในย่อหน้า: {full!r}")
    start, end = idx, idx + len(old)
    pos = 0
    replaced = False
    for r in runs:
        r_start, r_end = pos, pos + len(r.text)
        pos = r_end
        if r_end <= start or r_start >= end:
            continue
        head = r.text[:max(0, start - r_start)]
        tail = r.text[max(0, end - r_start):] if r_end > end else ""
        if not replaced:
            r.text = head + new + tail
            replaced = True
        else:
            r.text = head + tail


def build_drug_test_115():
    doc = Document(FORMS + "drug_test_115_original.docx")
    p = doc.paragraphs

    # (ดัชนีย่อหน้า, ข้อความเดิมในฟอร์ม, แท็กที่ใส่แทน)
    subs = [
        (3, "สถานีตำรวจภูธรนายายอาม", "{station_name}"),
        (4, "16 เดือน กันยายน พ.ศ. 2569", "{record_day} เดือน {record_month} พ.ศ. {record_year}"),
        (5, "ร.ต.อ.รัฐภูมิ พวงมาลา", "{officer1_rank}{officer1_name}"),
        (5, "รอง สว.สส.สภ.นายายอาม", "{officer1_position}"),
        (6, "ผบก.ภ.จว.จันทบุรี", "{officer_affiliation}"),
        (7, "☑", "{cb_card_ppst}"),
        (7, "6705928", "{card_no_ppst}"),
        (8, "☐", "{cb_card_gov}"),
        (8, "เลขที่", "เลขที่ {card_no_gov}"),
        (10, "นายวรากร วรรณมาลี", "{suspect_title}{suspect_full_name}"),
        (10, "อายุ 36 ปี", "อายุ {suspect_age} ปี"),
        (11, "☑ บัตรประชาชน", "{cb_id_card} บัตรประชาชน"),
        (11, "☐ บัตรคนซึ่งไม่มีสัญชาติไทย", "{cb_id_alien} บัตรคนซึ่งไม่มีสัญชาติไทย"),
        (11, "☐ หนังสือเดินทาง", "{cb_id_passport} หนังสือเดินทาง"),
        (11, "☐ เอกสารอื่นที่ราชการออกให้ ระบุ", "{cb_id_other} เอกสารอื่นที่ราชการออกให้ ระบุ {id_other_detail}"),
        (12, "1 2299 00306 18 6", "{suspect_id_number}"),
        (13, "หมู่ที่ 4 ตำบลช้างข้าม อำเภอนายายอาม จังหวัดจันทบุรี", "{suspect_address}"),
        (14, "หมู่ที่ 4 ตำบลช้างข้าม อำเภอนายายอาม จังหวัดจันทบุรี", "{suspect_current_address}"),
        (14, "หมายเลขโทรศัพท์", "หมายเลขโทรศัพท์ {suspect_phone}"),
        (16, "☑", "{cb_test_pos}"),
        (16, "เมทแอมเฟตามีน", "{drug_type}"),
        (17, "☐", "{cb_test_neg}"),
        (19, "☑", "{cb_search_none}"),
        (20, "☐", "{cb_search_found}"),
        (20, "ประเภท/ชนิด  ปริมาณ", "ประเภท/ชนิด {search_drug_type} ปริมาณ {search_amount}"),
        (24, "☐", "{cb_pb1}"),
        (25, "☐", "{cb_pb2}"),
        (26, "☐", "{cb_pb3}"),
        (27, "☑", "{cb_pb4}"),
        (29, "อาชีพ", "อาชีพ {suspect_occupation}"),
        (29, "รายได้โดยประมาณ", "รายได้โดยประมาณ {suspect_income}"),
        (32, "วรากร วรรณมาลี", "{suspect_full_name}"),
        (33, "☐ ขอลงนามสมัครใจ", "{cb_consent_yes} ขอลงนามสมัครใจ"),
        (33, "☑ ขอลงนามไม่สมัครใจ", "{cb_consent_no} ขอลงนามไม่สมัครใจ"),
        (42, "ในวันที่  เวลา  น.", "ในวันที่ {appointment_date} เวลา {appointment_time} น."),
        (43, "ณ สถานที่", "ณ สถานที่ {appointment_place}"),
        (53, "วรากร วรรณมาลี", "{suspect_full_name}"),
        (55, "ร.ต.อ.", "{officer1_rank}"),
        (56, "รัฐภูมิ พวงมาลา", "{officer1_name}"),
        (59, "วรากร วรรณมาลี", "{suspect_full_name}"),
    ]
    for idx, old, new in subs:
        replace_across_runs(p[idx], old, new)

    doc.save(OUT + "drug_test_115_template.docx")
    print("built drug_test_115_template.docx from real form")


def build_urine_referral():
    doc = Document(FORMS + "urine_referral_original.docx")
    p = doc.paragraphs

    # [2] ที่ <เลขหนังสือ> --tab--> <ชื่อสถานี>
    # เลขที่หนังสือยาวไม่เท่ากันทุกวัน (ออกเลขจากระบบอื่นแล้วมากรอก) ฟอร์มเดิมใช้เคาะเว้นวรรค
    # จัดตำแหน่ง ชื่อสถานีจึงเลื่อนตามความยาวเลข เปลี่ยนมาใช้ tab stop ตายตัวแทน
    p[2].paragraph_format.tab_stops.add_tab_stop(LETTERHEAD_TAB, WD_TAB_ALIGNMENT.LEFT)
    set_runs(p[2], {3: "{doc_number}", 14: "{station_name}"})
    clear_runs(p[2], [4, 5, 6, 7, 8, 10, 11, 12, 13])  # เหลือ run9 ที่เป็นแท็บเดียว

    # [3] จังหวัด.... รหัสไปรษณีย์ — ใช้ tab stop เดียวกันให้ตรงกับชื่อสถานีเสมอ
    p[3].paragraph_format.tab_stops.add_tab_stop(LETTERHEAD_TAB, WD_TAB_ALIGNMENT.LEFT)
    set_runs(p[3], {0: "\t", 1: "จังหวัด{province}", 2: "  {postal_code}"})

    # [4] วันที่ออกหนังสือ
    set_runs(p[4], {6: "{doc_date}"})
    clear_runs(p[4], [7, 8, 9, 10])

    # [6] เรียน ผู้อำนวยการ<โรงพยาบาล>
    set_runs(p[6], {2: "ผู้อำนวยการ{hospital_name}"})

    # [7] ย่อหน้าเนื้อความ (ทุก run รูปแบบเดียวกัน ยุบเหลือ run เดียวได้)
    flatten(p[7],
            "เนื่องด้วย {station_name} ขอส่งตรวจยืนยัน{drug_type} สารเสพติดในปัสสาวะ "
            "ในขั้นที่สองด้วยหลักการทางวิทยาศาสตร์ เพื่อเป็นการดำเนินการตามแนวทางการตรวจพิสูจน์"
            "หาสารเสพติดในปัสสาวะตามพระราชบัญญัติฟื้นฟูสมรรถภาพผู้ติดยาเสพติด พ.ศ. ๒๕๔๕ "
            "ที่ห้องเคมีคลินิกและพิษวิทยา กลุ่มงานพยาธิวิทยาคลินิก {hospital_name} {province} "
            "ในวันที่ {test_date} เป็นจำนวน {test_count} ราย  ดังมีรายชื่อต่อไปนี้")

    # [9] รายชื่อผู้ถูกส่งตรวจ — เหลือย่อหน้าเดียวเป็นต้นแบบ ตอนสร้างเอกสารจะโคลนตามจำนวนคน
    flatten(p[9], "{suspect_list_block}")

    # [10] ตัวอย่างรายชื่อคนที่สองในฟอร์มเดิม ลบทิ้ง
    p[10]._element.getparent().remove(p[10]._element)

    # [18] บรรทัดยศเหนือวงเล็บ (ที่ว่างด้านขวาของยศไว้เซ็นชื่อ) — ต้องชิดซ้ายให้ตรงขอบซ้ายของวงเล็บ
    # บรรทัดชื่อจัดกึ่งกลาง ขอบซ้ายจึงขึ้นกับความยาวชื่อ → ให้แอปคำนวณระยะเยื้อง {signer_rank_indent} (twips) ตอน export
    rank_run = p[18].add_run("{signer_rank}")
    rank_run.font.name = p[19].runs[2].font.name
    rank_pPr = p[18]._p.get_or_add_pPr()
    rank_ind = rank_pPr.find(qn("w:ind"))
    rank_ind.attrib.pop(qn("w:firstLine"), None)
    rank_ind.set(qn("w:left"), "{signer_rank_indent}")
    rank_pPr.find(qn("w:jc")).set(qn("w:val"), "left")

    # [19] ( ชื่อผู้ลงนาม )  /  [20] ตำแหน่งผู้ลงนาม
    set_runs(p[19], {2: "{signer_name} "})
    clear_runs(p[19], [3, 4])
    flatten(p[20], "{signer_position}")

    # ตารางท้ายเอกสาร ฝั่ง "สำหรับการรับ-ส่งตัวอย่าง" — ผู้ส่ง/วันที่/เวลา ที่นำส่งตัวอย่าง
    rows = doc.tables[0].rows
    rows[1].cells[0].paragraphs[0].runs[0].text = "ผู้ส่ง {signer_rank}{signer_name}"
    rows[3].cells[0].paragraphs[0].runs[1].text = "  {sample_send_date}"
    rows[4].cells[0].paragraphs[0].runs[1].text = "   {sample_send_time}   "

    doc.save(OUT + "urine_referral_template.docx")
    print("built urine_referral_template.docx from real form")


# ใบปะหน้าส่งศาล: ขอบซ้ายกระดาษ 1701 twips (3 ซม.) บล็อกชื่อสถานีในฟอร์มจริงเริ่มประมาณ 10.9 ซม. จากขอบพิมพ์
COURT_LETTERHEAD_TAB = Cm(10.9)


def build_court_referral():
    doc = Document(FORMS + "court_referral_original.docx")
    p = doc.paragraphs

    # [0] ที่ <เลขหนังสือ> --tab--> <ชื่อสถานี>  (ฟอร์มเดิมใช้แท็บ 5 ตัว + เคาะวรรค เปลี่ยนเป็น tab stop ตายตัว)
    p[0].paragraph_format.tab_stops.add_tab_stop(COURT_LETTERHEAD_TAB, WD_TAB_ALIGNMENT.LEFT)
    set_runs(p[0], {3: "{court_doc_number}", 4: "", 5: "\t", 11: "{station_name}"})
    clear_runs(p[0], [6, 7, 8, 9, 10])

    # [1] จังหวัด.... รหัสไปรษณีย์ — tab stop เดียวกัน
    p[1].paragraph_format.tab_stops.add_tab_stop(COURT_LETTERHEAD_TAB, WD_TAB_ALIGNMENT.LEFT)
    set_runs(p[1], {0: "\t", 6: "จังหวัด{province}  {postal_code}"})
    clear_runs(p[1], [1, 2, 3, 4, 5])

    # [2] วันที่ออกหนังสือ (ฟอร์มนี้ไม่มีคำว่า พ.ศ.) คงแท็บ 6 ตัวเดิมไว้
    set_runs(p[2], {6: " {doc_date_no_era}"})
    clear_runs(p[2], [7, 8, 9, 10])

    # [3] เรื่อง / [4] เรียน — ชื่อศาลเปลี่ยนตามที่กรอก
    set_runs(p[3], {2: "ส่งตัวผู้ต้องหาตามหมายจับของ{court_name}"})
    set_runs(p[4], {3: "{court_name}"})

    # [5] อ้างถึง ... — run0 "อ้างถึง" เป็นตัวหนา (คงไว้) เนื้อความอยู่ run2 เป็นตัวปกติ
    # ฟอร์มเดิมพิมพ์ "สิ่งที่ส่งมาด้วย" ต่อท้ายในย่อหน้าเดียวกันโดยไม่ขึ้นบรรทัดใหม่ แยกให้ขึ้นบรรทัดใหม่
    # (ต้องเป็นย่อหน้าใหม่ ไม่ใช่ line break — ย่อหน้านี้จัดแบบ thaiDistribute ถ้าใช้ line break Word จะยืดตัวอักษรบรรทัดก่อนตัด)
    # {case_red_clause} = " คดีหมายเลขแดงที่ …" เฉพาะเมื่อหมายจับมีเลขแดง (บางหมายไม่มี) แอปเป็นคนประกอบให้
    set_runs(p[5], {2: "จับตามหมายจับของ{court_name} ที่ {warrant_no} คดีหมายเลขดำที่ {case_black_no}{case_red_clause} "
                       "ลงวันที่ {warrant_date} ในความผิดฐาน “{charge}”"})
    clear_runs(p[5], range(3, len(p[5].runs)))
    attach_p = copy.deepcopy(p[5]._p)
    p[5]._p.addnext(attach_p)
    attach = Paragraph(attach_p, p[5]._parent)
    set_runs(attach, {0: "สิ่งที่ส่งมาด้วย", 2: "{court_attachment}"})
    attach.paragraph_format.space_before = 0
    p = doc.paragraphs  # ดัชนีหลังจากนี้เลื่อนไป 1

    # [7] รายละเอียดหมายศาล (วันแรกคือวันที่ศาลออกหมาย = ลงวันที่หมายจับ)
    flatten(p[7],
            "ด้วยเมื่อวันที่ {warrant_date} {court_name} ที่ {warrant_no} คดีหมายเลขดำที่ {case_black_no}{case_red_clause} "
            "ลงวันที่ {warrant_date} ในความผิดฐาน {charge}นั้น")

    # [8] ผู้ต้องหา + ผู้นำตัวส่งศาล — ฟอร์มนี้ใช้ยศ/ตำแหน่งตัวเต็มทั้งหมด
    flatten(p[8],
            "{station_name} จังหวัด{province} ขอเรียนว่า ได้ทำการจับกุม {suspect_title}{suspect_name_with_nick} "
            "อายุ {suspect_age} ปี สัญชาติ {suspect_nationality} หมายเลขประจำตัวประชาชน {suspect_id_number} "
            "ที่อยู่ {suspect_address_court} จึงมอบหมายให้ {officer1_rank_full}{officer1_name} {officer1_position_full} "
            "{station_name} จังหวัด{province} พร้อมพวก เป็นผู้นำตัว{suspect_title}{suspect_name_with_nick} "
            "มาส่งตัวที่{court_name} เพื่อดำเนินการต่อไป พร้อมนี้ ได้แนบบันทึกการจับกุมตัว มาพร้อมนี้ด้วยแล้ว จำนวน ๑ ฉบับ")

    # ผู้ลงนาม (รอง ผกก. ปฏิบัติราชการแทน ผกก.) — คงการเคาะวรรคจัดตำแหน่งของฟอร์มเดิมไว้
    replace_across_runs(p[13], "พันตำรวจโท", "{court_signer_rank_full}")
    replace_across_runs(p[15], "วรรณวุฒิ แสนเสนยา", "{court_signer_name}")
    replace_across_runs(p[16], "รองผู้กำกับการสืบสวนปฏิบัติราชการแทน", "{court_signer_position1}")
    replace_across_runs(p[17], "ผู้กำกับการสถานีตำรวจภูธรนายายอาม", "{court_signer_position2}")

    # ท้ายกระดาษ
    flatten(p[22], "{station_name} จังหวัด{province}")

    doc.save(OUT + "court_referral_template.docx")
    print("built court_referral_template.docx from real form")


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    build_urine_referral()
    build_drug_test_115()
    build_court_referral()
