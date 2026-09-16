"""แปลงไฟล์ฟอร์มจริงของสถานี (forms/*.docx) ให้เป็นแม่แบบที่มีแท็ก {...} โดยไม่แตะเลย์เอาต์เดิม

ต่างจาก generate_templates.py ที่ "สร้างเอกสารขึ้นมาใหม่" ทั้งฉบับ — สคริปต์นี้เปิดไฟล์ฟอร์มจริง
แล้วแทนที่เฉพาะ "ข้อความในช่องที่ต้องกรอก" ด้วยแท็ก ทำให้ฟอนต์/ระยะ/ระยะบรรทัด/ตาราง
ของฟอร์มเดิมคงอยู่ครบ เวลา export ออกมาฟอร์มจึงไม่ขยับ
"""
import os
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.shared import Cm

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

    # [18] บรรทัดว่างเหนือชื่อผู้ลงนาม ใส่ยศของเจ้าหน้าที่ที่เลือกไว้ตอนล็อกอิน
    rank_run = p[18].add_run("{signer_rank}")
    rank_run.font.name = p[19].runs[2].font.name

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


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    build_urine_referral()
