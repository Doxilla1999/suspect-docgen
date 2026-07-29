// Bookmarklet: กรอกฟอร์มขั้นตอนที่ 1 ของ arrest.dopa.go.th (หน้า "แจ้งการควบคุมตัว")
// อัตโนมัติจากข้อมูลที่คัดลอกไว้จากเว็บแอป suspect-docgen (ปุ่ม "คัดลอกข้อมูลสำหรับ DOPA")
//
// ไม่ได้กดบันทึก/ส่งฟอร์มให้ — แค่กรอกช่องให้ ต้องตรวจสอบและกดปุ่ม "เพิ่มสถานที่" เองเสมอ
//
// วิธีใช้: ดู README.md ในโฟลเดอร์นี้

(async function () {
  function setVal(id, val) {
    const el = document.getElementById(id);
    if (!el || !val) return false;
    el.value = val;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  function selectByText(selectEl, text) {
    if (!selectEl || !text) return false;
    const opt = [...selectEl.options].find(o => o.textContent.trim() === text.trim());
    if (!opt) return false;
    selectEl.value = opt.value;
    selectEl.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  function waitForOptions(selectEl, timeoutMs) {
    return new Promise(resolve => {
      if (selectEl.options.length > 1) return resolve(true);
      const obs = new MutationObserver(() => {
        if (selectEl.options.length > 1) { obs.disconnect(); resolve(true); }
      });
      obs.observe(selectEl, { childList: true });
      setTimeout(() => { obs.disconnect(); resolve(selectEl.options.length > 1); }, timeoutMs);
    });
  }

  let data;
  try {
    data = JSON.parse(await navigator.clipboard.readText());
  } catch (e) {
    alert('อ่านข้อมูลจาก clipboard ไม่ได้ — กด "คัดลอกข้อมูลสำหรับ DOPA" ในเว็บแอป suspect-docgen ก่อน แล้วค่อยกด bookmarklet นี้');
    return;
  }
  if (!data || data.source !== 'suspect-docgen-dopa-step1') {
    alert('ข้อมูลใน clipboard ไม่ใช่ข้อมูลจาก suspect-docgen — คัดลอกใหม่อีกครั้ง');
    return;
  }

  const notes = [];

  setVal('place_name', data.place_name);
  setVal('place_date', data.place_date);

  const pSel = document.getElementById('place_pcode');
  if (pSel && data.province_code) {
    const ok = selectByText(pSel, data.province_name) || (pSel.value = data.province_code, pSel.dispatchEvent(new Event('change', { bubbles: true })), true);
    if (ok && typeof dochange === 'function') {
      dochange('amphoe', data.province_code);
      const aSel = document.getElementById('place_acode');
      if (aSel) {
        await waitForOptions(aSel, 4000);
        if (selectByText(aSel, data.amphoe_name)) {
          // อำเภอเลือกได้ — คลื่นถัดไป (ตำบล) โหลดผ่านกลไกของหน้าเว็บ DOPA เอง
          // เฝ้าดู #place_tcode สักพัก เผื่อระบบเติม option ให้อัตโนมัติหลัง onchange ของ อำเภอ
          const tSel = document.getElementById('place_tcode');
          if (tSel) {
            await waitForOptions(tSel, 4000);
            if (!selectByText(tSel, data.tambon_name)) {
              notes.push('เลือก "ตำบล" เองด้วยนะครับ (ระบบไม่เติมตัวเลือกให้อัตโนมัติ หรือหาชื่อตำบลไม่เจอในรายการ)');
            }
          }
        } else {
          notes.push('เลือก "อำเภอ" และ "ตำบล" เองด้วยนะครับ (หาชื่ออำเภอไม่เจอในรายการที่โหลดมา)');
        }
      }
    } else {
      notes.push('เลือก "อำเภอ" และ "ตำบล" เองด้วยนะครับ');
    }
  }

  let msg = 'กรอกข้อมูลที่ทำได้อัตโนมัติแล้ว — ตรวจสอบทุกช่องก่อนกด "เพิ่มสถานที่" เสมอ';
  if (notes.length) msg += '\n\n' + notes.join('\n');
  alert(msg);
})();
