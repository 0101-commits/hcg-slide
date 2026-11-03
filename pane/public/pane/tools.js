/* 「HCG 장표」 3단계 — 부품 · 아이콘 · 칸 탭. pane.js 의 window.HCGPane 과 slotops.js 의 window.HCGSlot 을 씀 */
(() => {
  "use strict";
  const $ = (s) => document.querySelector(s);
  const H = () => window.HCGPane;
  const EMU = 12700; // 1pt
  const today = () => new Date().toISOString().slice(0, 10);

  async function copyText(t) {
    try { await navigator.clipboard.writeText(t); return true; } catch (_) { /* 클립보드 막힘 → 옛 방식 */ }
    const ta = document.createElement("textarea");
    ta.value = t; document.body.appendChild(ta); ta.select();
    let ok = false;
    try { ok = document.execCommand("copy"); } catch (_) { ok = false; }
    ta.remove();
    return ok;
  }

  // 지금 파일에 PPTX(base64) 장을 넣고 새 장 번호를 돌려줌
  async function insertB64(b64, fmt, tags, afterId) {
    let at = 0, newId = null;
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items/id");
      await ctx.sync();
      const before = new Set(slides.items.map((s) => s.id));
      const opts = { formatting: fmt };
      const target = afterId || (sel.items.length ? sel.items[sel.items.length - 1].id : null);
      if (target) opts.targetSlideId = target;
      ctx.presentation.insertSlidesFromBase64(b64, opts);
      await ctx.sync();
      const after = ctx.presentation.slides;
      after.load("items/id");
      await ctx.sync();
      const added = after.items.filter((s) => !before.has(s.id));
      for (const s of added) for (const [k, v] of Object.entries(tags || {})) s.tags.add(k, String(v).slice(0, 250));
      await ctx.sync();
      if (added.length) {
        newId = added[0].id;
        at = after.items.findIndex((s) => s.id === newId) + 1;
        ctx.presentation.setSelectedSlides([newId]);
        await ctx.sync();
      }
    });
    return { at, newId };
  }

  // ================= 부품 =================
  let parts = null;
  async function loadParts() {
    if (parts) return;
    const r = await fetch("parts.json", { credentials: "same-origin" });
    parts = await r.json();
    const grid = $("#parts");
    grid.replaceChildren(...parts.map((p) => {
      const b = document.createElement("button");
      b.className = "card"; b.type = "button";
      const v = p.variants[0];
      b.innerHTML = `<img loading="lazy" alt="" src="/api/thumb/${v.id}.jpg"><div class="t">${H().esc(p.name)}</div>
        <div class="m">${p.variants.length}가지 모양 · 템플릿 ${p.count}장</div>`;
      b.onclick = () => openPart(p);
      return b;
    }));
  }
  function openPart(p) {
    $("#parts").hidden = true;
    $("#part-view").hidden = false;
    $("#part-title").textContent = p.name;
    $("#part-src").textContent = `구조 ${p.struct} · HCG 템플릿에서 고른 모양 ${p.variants.length}가지`;
    $("#part-vars").replaceChildren(...p.variants.map((v) => {
      const d = document.createElement("div");
      d.className = "card";
      d.innerHTML = `<img alt="" src="/api/img/${v.id}.jpg"><div class="t">${H().esc(v.title || p.name)}</div>
        <div class="m">${H().esc(v.doc)}</div>`;
      const btn = document.createElement("button");
      btn.className = "primary"; btn.type = "button"; btn.textContent = "이 모양 넣기";
      btn.style.margin = "0 7px 7px";
      btn.onclick = () => insertPart(p, v, btn);
      d.appendChild(btn);
      return d;
    }));
  }
  async function insertPart(p, v, btn) {
    if (!H().inOffice()) { H().toast("PowerPoint 안에서 넣을 수 있음", true); return; }
    btn.disabled = true; btn.textContent = "넣는 중…";
    try {
      const buf = await H().api(`slide/${v.id}.pptx`, "buf");
      const fmt = document.querySelector('input[name="pfmt"]:checked').value;
      const { at } = await insertB64(H().toBase64(buf), fmt, { HCG_PART: p.code, HCG_SRC: v.id, HCG_ADDED: today() });
      H().toast(`${at}번 장에 「${p.name}」 부품을 넣음 — 칸 수는 「칸」 탭에서`);
    } catch (e) {
      if (e && e.message !== "login") H().toast("넣지 못함 — " + (e.message || e), true);
    } finally { btn.disabled = false; btn.textContent = "이 모양 넣기"; }
  }

  // ================= 아이콘 =================
  let icons = null, iconShown = 0, iconList = [];
  async function loadIcons() {
    if (icons) return;
    icons = await (await fetch("icons.json", { credentials: "same-origin" })).json();
    $("#icon-count").textContent = `${icons.length}개`;
    filterIcons();
  }
  function filterIcons() {
    const terms = $("#iq").value.trim().split(/\s+/).filter(Boolean);
    iconList = icons.filter((i) => terms.every((t) => (i.name + " " + i.words).includes(t)));
    iconShown = 0;
    $("#icons").replaceChildren();
    moreIcons();
  }
  function moreIcons() {
    const next = iconList.slice(iconShown, iconShown + 60);
    $("#icons").append(...next.map((i) => {
      const b = document.createElement("button");
      b.type = "button"; b.title = `${i.id} · ${i.name}\n${i.words}`;
      b.innerHTML = `<img loading="lazy" alt="${H().esc(i.name)}" src="icons/${i.id}.svg">`;
      b.onclick = () => insertIcon(i);
      return b;
    }));
    iconShown += next.length;
    $("#icon-more").hidden = iconShown >= iconList.length;
    $("#icon-more").textContent = `더 보기 (${iconShown} / ${iconList.length})`;
  }
  async function insertIcon(i) {
    if (!H().inOffice()) { H().toast("PowerPoint 안에서 넣을 수 있음", true); return; }
    const svg = await (await fetch(`icons/${i.id}.svg`, { credentials: "same-origin" })).text();
    Office.context.document.setSelectedDataAsync(svg, { coercionType: Office.CoercionType.XmlSvg, imageWidth: 48, imageHeight: 48 }, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) H().toast(`아이콘 「${i.name}」을 넣음`);
      else H().toast("아이콘을 넣지 못함 — " + (res.error && res.error.message), true);
    });
  }

  // ================= 칸 =================
  const cur = { slideId: null, no: 0, zip: null, xmlName: null, xml: null, slots: [], cols: [], tables: [] };
  const ROLE = { title: "제목", key: "핵심 문장", head: "머리", body: "본문", label: "꼬리표", note: "각주" };
  let reading = false, again = false, timer = 0;

  async function readSlide() {
    if (!H().inOffice()) { $("#slot-title").textContent = "PowerPoint 안에서 열어야 함"; return; }
    if (!window.HCGSlot) { $("#slot-title").textContent = "칸 엔진을 불러오지 못함"; return; }
    if (reading) { again = true; return; }
    reading = true;
    try {
      let b64 = null, slideId = null, no = 0;
      await PowerPoint.run(async (ctx) => {
        const all = ctx.presentation.slides;
        all.load("items/id");
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        if (!sel.items.length) return;
        slideId = sel.items[0].id;
        no = all.items.findIndex((s) => s.id === slideId) + 1;
        const r = sel.items[0].exportAsBase64();
        await ctx.sync();
        b64 = r.value;
      });
      if (!b64) { $("#slot-title").textContent = "왼쪽에서 장을 하나 고르세요"; $("#slots").replaceChildren(); return; }
      const zip = await JSZip.loadAsync(b64, { base64: true });
      const xmlName = Object.keys(zip.files).find((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
      const xml = await zip.file(xmlName).async("string");
      const pres = zip.file("ppt/presentation.xml");
      const opts = pres ? HCGSlot.slideSize(await pres.async("string")) : {};
      Object.assign(cur, { slideId, no, zip, xmlName, xml, opts });
      cur.slots = HCGSlot.slotMap(xml, opts).filter((s) => !/^HCG_NUM_/.test(s.name || ""));
      cur.cols = HCGSlot.columns(xml, opts) || [];
      const ov = HCGSlot.overlaps(cur.slots) || [];
      cur.ovSet = new Set(ov.flatMap((p) => [p[0] && p[0].no, p[1] && p[1].no, p.a, p.b].filter((x) => x != null)));
      renderSlots(ov.length);
    } catch (e) {
      console.error(e);
      $("#slot-title").textContent = "장을 읽지 못함 — " + (e.message || e);
    } finally {
      reading = false;
      if (again) { again = false; readSlide(); }
    }
  }

  function renderSlots(nOverlap) {
    const texts = cur.slots.filter((s) => s.chars > 0);
    const over = texts.filter((s) => s.overflow).length;
    $("#slot-title").textContent = `${cur.no}번 장 · 글 칸 ${texts.length}개`;
    $("#slot-sum").textContent = `열 ${cur.cols.length || "—"} · 넘침 어림 ${over} · 겹침 ${nOverlap}`;
    $("#slot-warn").textContent = over || nOverlap ? "주황 표시 칸을 확인하세요" : "";
    // 열
    $("#slot-cols").hidden = cur.cols.length < 2;
    $("#col-n").textContent = cur.cols.length;
    $("#col-minus").disabled = cur.cols.length <= 2;
    $("#col-plus").disabled = cur.cols.length >= 5;
    // 표
    const tables = [...new Set(cur.slots.filter((s) => s.kind === "cell").map((s) => s.tableNo))];
    cur.tables = tables;
    const tw = $("#slot-tables");
    tw.hidden = !tables.length;
    if (tables.length) {
      const rows = (t) => Math.max(...cur.slots.filter((s) => s.kind === "cell" && s.tableNo === t).map((s) => s.r)) + 1;
      tw.innerHTML = `<span>표 행 지우기</span><select id="tb-sel">${tables.map((t, i) => `<option value="${t}">표 ${i + 1} (${rows(t)}행)</option>`).join("")}</select>
        <select id="row-sel"></select><button class="ghost sm" id="row-del" type="button">지우기</button>`;
      const fillRows = () => {
        const t = +$("#tb-sel").value;
        $("#row-sel").innerHTML = Array.from({ length: rows(t) }, (_, i) => `<option value="${i}">${i + 1}행${i === 0 ? "(머리)" : ""}</option>`).join("");
        if (rows(t) > 1) $("#row-sel").value = "1";
      };
      $("#tb-sel").onchange = fillRows;
      fillRows();
      $("#row-del").onclick = () => applyOp("row", () => HCGSlot.deleteTableRow(cur.xml, +$("#tb-sel").value, +$("#row-sel").value, cur.opts), `표 ${$("#tb-sel").selectedIndex + 1} 의 ${+$("#row-sel").value + 1}행 지움`);
    }
    // 칸 목록
    $("#slots").replaceChildren(...texts.map((s) => {
      const li = document.createElement("li");
      const bad = s.overflow || cur.ovSet.has(s.no) || s.wchars > s.limitHi;
      if (bad) li.className = "bad";
      const pct = Math.min(100, Math.round((s.wchars / Math.max(1, s.limitHi)) * 100));
      const where = s.kind === "cell" ? ` · 표 ${cur.tables.indexOf(s.tableNo) + 1} ${s.r + 1}행 ${s.c + 1}열` : "";
      li.innerHTML = `<div class="row1"><span class="no">${s.no}</span><span class="role">${ROLE[s.role] || s.role || ""}${where}</span>
        <span class="cnt">${Math.round(s.wchars)}자 · 한도 ${Math.round(s.limitLo)}~${Math.round(s.limitHi)}</span></div>
        <div class="txt">${H().esc(s.text)}</div><div class="bar"><i style="width:${pct}%"></i></div>
        <div class="flags">${s.overflow ? "<span>넘침 어림</span>" : ""}${cur.ovSet.has(s.no) ? "<span>겹침</span>" : ""}<button class="ask" type="button">Claude에게 ›</button></div>`;
      li.onclick = (ev) => { if (!ev.target.classList.contains("ask")) selectShape(s); };
      li.querySelector(".ask").onclick = async () => {
        const msg = `/hcg-restructure ${cur.no}번 장 ${s.no}번 칸(${ROLE[s.role] || "글"}, 「${(s.text || "").slice(0, 20)}…」)을 ${Math.round(s.limitLo)}~${Math.round(s.limitHi)}자(폭 가중) 안으로 다시 써 줘. 칸 · 서식은 그대로 두고 글만.`;
        H().toast((await copyText(msg)) ? "Claude 작업창에 붙여넣기(Ctrl+V) 하세요" : "복사하지 못함 — " + msg, false);
      };
      return li;
    }));
  }

  async function selectShape(s) {
    try {
      await PowerPoint.run(async (ctx) => {
        const slide = ctx.presentation.slides.getItem(cur.slideId);
        const sh = slide.shapes;
        sh.load("items/id,items/name,items/left,items/top,items/width,items/height");
        await ctx.sync();
        const cx = (s.x + s.w / 2) / EMU, cy = (s.y + s.h / 2) / EMU;
        // 같은 이름 중 가장 가까운 것, 없으면 칸을 감싸는 도형(묶음 · 표)
        let best = null, bd = 1e12;
        for (const it of sh.items) {
          const inside = cx >= it.left && cx <= it.left + it.width && cy >= it.top && cy <= it.top + it.height;
          const d = Math.hypot(it.left + it.width / 2 - cx, it.top + it.height / 2 - cy) + (it.name === s.name ? 0 : inside ? 500 : 5000);
          if (d < bd) { bd = d; best = it; }
        }
        if (best) { slide.setSelectedShapes([best.id]); await ctx.sync(); }
      });
    } catch (e) { console.warn(e); }
  }

  async function setNums(on) {
    if (!H().inOffice() || !cur.slideId) return;
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(cur.slideId);
      const sh = slide.shapes;
      sh.load("items/name");
      await ctx.sync();
      for (const it of sh.items) if (/^HCG_NUM_/.test(it.name)) it.delete();
      await ctx.sync();
      if (!on) return;
      for (const s of cur.slots.filter((x) => x.chars > 0)) {
        const tb = slide.shapes.addTextBox(String(s.no), { left: s.x / EMU, top: s.y / EMU, width: 20, height: 14 });
        tb.name = `HCG_NUM_${s.no}`;
        tb.fill.setSolidColor("#E8743B");
        tb.textFrame.textRange.font.size = 8;
        tb.textFrame.textRange.font.color = "#FFFFFF";
        tb.textFrame.textRange.font.bold = true;
        tb.textFrame.leftMargin = 1; tb.textFrame.rightMargin = 1; tb.textFrame.topMargin = 0; tb.textFrame.bottomMargin = 0;
      }
      await ctx.sync();
    });
  }

  // 장을 꺼내 고친 XML 로 다시 넣음
  async function applyOp(kind, fn, label) {
    try {
      if ($("#show-nums").checked) { $("#show-nums").checked = false; await setNums(false); await readSlide(); }
      const res = fn();
      if (!res || !res.xml || res.xml === cur.xml) throw new Error((res && res.warnings && res.warnings.join(" · ")) || "바꾸지 못함");
      cur.zip.file(cur.xmlName, res.xml);
      const b64 = await cur.zip.generateAsync({ type: "base64", compression: "DEFLATE" });
      const origId = cur.slideId;
      const { at, newId } = await insertB64(b64, "KeepSourceFormatting", { HCG_EDIT: kind, HCG_EDITED: today() }, origId);
      if (!$("#keep-orig").checked && newId) {
        await PowerPoint.run(async (ctx) => { ctx.presentation.slides.getItem(origId).delete(); await ctx.sync(); });
      }
      let warn = (res.warnings || []).length ? ` · 확인: ${res.warnings.slice(0, 2).join(" · ")}` : "";
      if (label.includes("→") && /열 (\d)→(\d)/.test(label) && +RegExp.$2 > +RegExp.$1) warn += " · 열이 좁아져 글이 넘칠 수 있음 — 주황 칸을 Claude에게 줄여 달라고 하세요";
      H().toast(`${label} → ${$("#keep-orig").checked ? at + "번 장에 새로" : "바꿈"}${warn}`);
      setTimeout(readSlide, 400);
    } catch (e) {
      console.error(e);
      H().toast("바꾸지 못함 — " + (e.message || e), true);
    }
  }

  function wire() {
    $("#part-back").onclick = () => { $("#part-view").hidden = true; $("#parts").hidden = false; };
    $("#iq").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(filterIcons, 200); });
    $("#icon-more").onclick = moreIcons;
    $("#slot-read").onclick = readSlide;
    $("#show-nums").onchange = (e) => setNums(e.target.checked).catch((er) => H().toast("번호를 못 붙임 — " + er.message, true));
    $("#col-minus").onclick = () => applyOp("cols", () => HCGSlot.setColumns(cur.xml, cur.cols.length - 1, cur.opts), `열 ${cur.cols.length}→${cur.cols.length - 1}`);
    $("#col-plus").onclick = () => applyOp("cols", () => HCGSlot.setColumns(cur.xml, cur.cols.length + 1, cur.opts), `열 ${cur.cols.length}→${cur.cols.length + 1}`);
    document.addEventListener("hcg-tab", (e) => {
      if (e.detail === "parts") { loadParts().catch(console.error); loadIcons().catch(console.error); }
      if (e.detail === "slots") readSlide();
    });
    try {
      Office.onReady(() => {
        try {
          Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
            if (!$("#tab-slots").hidden) { clearTimeout(cur.t); cur.t = setTimeout(readSlide, 600); }
          });
        } catch (_) { /* 선택 바뀜 알림 없음 */ }
      });
    } catch (_) { /* Office 밖 */ }
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wire); else wire();
})();
