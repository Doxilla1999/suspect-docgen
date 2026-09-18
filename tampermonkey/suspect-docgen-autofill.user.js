// ==UserScript==
// @name         suspect-docgen auto-fill (DOPA + e-Saraban)
// @namespace    suspect-docgen
// @version      1.2
// @description  กรอกข้อมูลอัตโนมัติจากเว็บแอป suspect-docgen ลงเว็บ arrest.dopa.go.th และ saraban.police.go.th (ไม่กดบันทึก/ส่งให้)
// @match        https://arrest.dopa.go.th/*
// @match        https://saraban.police.go.th/saraban/*
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  function setVal(id, val) {
    const el = document.getElementById(id);
    if (!el || !val) return false;
    // ใช้ native setter ตรงๆ (ไม่ผ่าน el.value ปกติ) เพราะบางฟิลด์ของเว็บนี้ (เช่น place_date)
    // ถูกควบคุมด้วย framework ฝั่งหน้าเว็บ ซึ่งจะเซ็ตค่าทับกลับถ้าไม่ทำแบบนี้
    const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
    const nativeSetter = Object.getOwnPropertyDescriptor(proto, 'value') && Object.getOwnPropertyDescriptor(proto, 'value').set;
    if (nativeSetter) nativeSetter.call(el, val); else el.value = val;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('blur', { bubbles: true }));
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

    // 1.2 เจ้าหน้าที่ของรัฐผู้รับผิดชอบ (ม.22) — คนเดิมทุกครั้ง (ไม่ใช่คนเดียวกับข้อ 1.1)
    const RESPONSIBLE_OFFICER = {
      pid: '3560300569254',
      title: 'ร.ต.อ.',
      fname: 'รัฐภูมิ',
      lname: 'พวงมาลา',
      position: 'รอง สว.สส.',
      org: 'สภ.นายายอาม',
      tel: '0861801202'
    };
    const otherRadio = document.getElementById('other');
    if (otherRadio) {
      otherRadio.checked = true;
      otherRadio.dispatchEvent(new Event('change', { bubbles: true }));
      if (typeof check1_1 === 'function') check1_1(otherRadio);
      setVal('com_pid', RESPONSIBLE_OFFICER.pid);
      setVal('com_title', RESPONSIBLE_OFFICER.title);
      setVal('com_fname', RESPONSIBLE_OFFICER.fname);
      setVal('com_lname', RESPONSIBLE_OFFICER.lname);
      setVal('com_position', RESPONSIBLE_OFFICER.position);
      setVal('com_org', RESPONSIBLE_OFFICER.org);
      setVal('com_tel', RESPONSIBLE_OFFICER.tel);
    } else {
      notes.push('กรอก "1.2 เจ้าหน้าที่ของรัฐผู้รับผิดชอบ (ม.22)" เองด้วยนะครับ (หาปุ่มตัวเลือกไม่เจอ)');
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

  // ---------- e-Saraban: หน้า "สร้าง/หนังสือส่งภายใน" (ฟอร์ม #adddocform) ----------
  // ฟอร์มนี้ใช้ jQuery ธรรมดา เซ็ตค่าลง input ตรงๆ ได้
  // ไม่แตะ: เลขที่เอกสาร (txtwid/txtregno) ระบบ/ผู้ใช้จัดการเอง, วันที่ (ระบบเติมวันปัจจุบันให้แล้ว),
  //         dropdown ชั้นความเร็ว/ความลับ/หมวดหนังสือ (ค่าเฉพาะหน่วยงาน) และไม่กดปุ่ม "สร้าง"
  // หมวดหนังสือสำหรับส่งตรวจปัสสาวะ = "หนังสือส่งภายนอก" ตัวแรกถัดจาก "หนังสือส่งภายใน" ใน dropdown
  // (รหัสเป็นค่าที่หน่วยตั้งเอง ถ้ารายการเปลี่ยนจะไม่เจอ สคริปต์จะให้เลือกเองแทน)
  const SARABAN_BOOKGROUP_EXTERNAL = '0043';

  function selectBookGroup(value) {
    const sel = document.getElementById('selbookgroup');
    if (!sel) return false;
    const opt = [...sel.options].find(o => o.value === value);
    if (!opt || !opt.textContent.includes('ภายนอก')) return false;
    // เป็น select2 — ผ่าน jQuery ถ้ามี เพื่อให้ตัว dropdown และ handler ของหน้าเว็บอัปเดตเหมือนคนเลือกเอง
    if (window.jQuery) window.jQuery(sel).val(value).trigger('change');
    else { sel.value = value; sel.dispatchEvent(new Event('change', { bubbles: true })); }
    return true;
  }

  async function fillSarabanUrine(data) {
    if (!document.getElementById('adddocform')) {
      alert('ยังไม่เจอฟอร์มสร้างหนังสือ — เปิดเมนู ลงทะเบียนรับส่ง แล้วเข้าหน้าสร้างหนังสือก่อน แล้วค่อยกดปุ่มนี้');
      return;
    }
    // เลือกหมวดหนังสือก่อน แล้วรอให้หน้าเว็บจัดการ (บางกรณีเปลี่ยนหมวดแล้วฟอร์มรีเซ็ต) ค่อยกรอกข้อความ
    const bookGroupOk = selectBookGroup(SARABAN_BOOKGROUP_EXTERNAL);
    if (!bookGroupOk) {
      alert('หาหมวด "หนังสือส่งภายนอก" ในรายการไม่เจอ — เลือกหมวดหนังสือเองก่อน แล้วกดปุ่มนี้อีกครั้ง');
      return;
    }
    await new Promise(r => setTimeout(r, 800));

    const fromEl = document.getElementById('txtfrom');
    if (fromEl && !fromEl.value.trim()) setVal('txtfrom', data.from);
    setVal('txtto', data.to);
    setVal('txtwsubject', data.subject);
    setVal('txtwdsc', data.detail);
    alert('เลือกหมวดหนังสือส่งภายนอก และกรอก จาก/ถึง/เรื่อง/รายละเอียด ให้แล้ว\n\nตรวจสอบทุกช่อง แล้วกด "สร้าง" เองนะครับ');
  }

  async function handleClick() {
    let data;
    try {
      data = JSON.parse(await navigator.clipboard.readText());
    } catch (e) {
      alert('อ่านข้อมูลจาก clipboard ไม่ได้ — กดปุ่ม "คัดลอกข้อมูล..." ในเว็บแอป suspect-docgen ก่อน แล้วค่อยกดปุ่มนี้');
      return;
    }
    if (!data || !data.source) {
      alert('ข้อมูลใน clipboard ไม่ใช่ข้อมูลจาก suspect-docgen — คัดลอกใหม่อีกครั้ง');
      return;
    }

    const onSaraban = location.hostname === 'saraban.police.go.th';
    if (data.source === 'suspect-docgen-saraban-urine' && onSaraban) {
      await fillSarabanUrine(data);
    } else if (data.source === 'suspect-docgen-dopa-step1' && location.pathname.includes('/detainee/')) {
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

  if (location.hostname === 'saraban.police.go.th') {
    // e-Saraban เป็นเว็บหน้าเดียว (URL ไม่เปลี่ยนตามเมนู) ฟอร์มถูกโหลดเข้ามาทีหลัง
    // จึงเช็คเป็นระยะว่าฟอร์มสร้างหนังสือโผล่มาหรือยัง แล้วค่อยโชว์/ซ่อนปุ่ม
    injectButton();
    const btn = document.getElementById('sdg-autofill-btn');
    setInterval(() => {
      const form = document.getElementById('adddocform');
      btn.style.display = (form && form.offsetParent !== null) ? '' : 'none';
    }, 1000);
  } else if (location.pathname.includes('/detainee/') || location.pathname.includes('/update_place/')) {
    injectButton();
  }
})();
