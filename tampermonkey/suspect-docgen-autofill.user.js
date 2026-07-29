// ==UserScript==
// @name         suspect-docgen auto-fill (DOPA)
// @namespace    suspect-docgen
// @version      1.0
// @description  กรอกข้อมูลอัตโนมัติจากเว็บแอป suspect-docgen ลงเว็บ arrest.dopa.go.th (ไม่กดบันทึก/ส่งให้)
// @match        https://arrest.dopa.go.th/*
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

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

  // ---------- หน้า detainee/ ("แจ้งการควบคุมตัว" — ขั้นตอนที่ 1) ----------
  async function fillDopaStep1(data) {
    const notes = [];

    setVal('place_name', data.place_name);
    setVal('place_date', data.place_date);

    const pSel = document.getElementById('place_pcode');
    if (pSel && data.province_code) {
      const ok = selectByText(pSel, data.province_name) ||
        (pSel.value = data.province_code, pSel.dispatchEvent(new Event('change', { bubbles: true })), true);
      if (ok && typeof dochange === 'function') {
        dochange('amphoe', data.province_code);
        const aSel = document.getElementById('place_acode');
        if (aSel) {
          await waitForOptions(aSel, 4000);
          if (selectByText(aSel, data.amphoe_name)) {
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
  }

  // ---------- หน้า update_place/ ("เพิ่มผู้ถูกควบคุมตัว" — ขั้นตอนที่ 2) ----------
  // ยังไม่ได้ทำ — ต้องได้ตัวอย่างหน้านี้จริงมาดูโครงสร้างฟอร์มก่อน (fname[]/lname[]/accusation[]/behavior[] ฯลฯ
  // เป็น field ที่เพิ่มแบบ dynamic ต่อคน ไม่มี id ตายตัว ต้องดูของจริงถึงจะกรอกให้ถูกต้องปลอดภัย)
  async function fillDopaStep2(data) {
    alert('ยังไม่รองรับหน้า "เพิ่มผู้ถูกควบคุมตัว" — รอการอัปเดตสคริปต์นี้');
  }

  async function handleClick() {
    let data;
    try {
      data = JSON.parse(await navigator.clipboard.readText());
    } catch (e) {
      alert('อ่านข้อมูลจาก clipboard ไม่ได้ — กด "คัดลอกข้อมูลสำหรับ DOPA" ในเว็บแอป suspect-docgen ก่อน แล้วค่อยกดปุ่มนี้');
      return;
    }
    if (!data || !data.source) {
      alert('ข้อมูลใน clipboard ไม่ใช่ข้อมูลจาก suspect-docgen — คัดลอกใหม่อีกครั้ง');
      return;
    }

    if (data.source === 'suspect-docgen-dopa-step1' && location.pathname.includes('/detainee/')) {
      await fillDopaStep1(data);
    } else if (data.source === 'suspect-docgen-dopa-step2' && location.pathname.includes('/update_place/')) {
      await fillDopaStep2(data);
    } else {
      alert('ข้อมูลใน clipboard ไม่ตรงกับหน้าเว็บนี้ (คัดลอกข้อมูลให้ตรงขั้นตอนก่อนนะครับ)');
    }
  }

  function injectButton() {
    if (document.getElementById('sdg-autofill-btn')) return;
    const btn = document.createElement('button');
    btn.id = 'sdg-autofill-btn';
    btn.textContent = '📋 กรอกจาก suspect-docgen';
    btn.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:99999;background:#B68D40;color:#201808;' +
      'border:none;border-radius:8px;padding:12px 18px;font-size:14px;font-weight:600;cursor:pointer;' +
      'box-shadow:0 2px 10px rgba(0,0,0,.3);font-family:sans-serif';
    btn.addEventListener('click', handleClick);
    document.body.appendChild(btn);
  }

  if (location.pathname.includes('/detainee/') || location.pathname.includes('/update_place/')) {
    injectButton();
  }
})();
