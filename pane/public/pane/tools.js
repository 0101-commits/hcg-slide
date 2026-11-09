/* 「HCG 장표」 Component · Item · Icon · Revision 탭. pane.js 의 window.HCGPane 과 slotops.js 의 window.HCGSlot 을 씀 */
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

  // ================= Item — 도형 하나 · 작은 묶음을 지금 장에 그대로 끼워 넣음 =================
  // 방법: 지금 장을 파일로 꺼냄 → 도형 XML 을 끼움 → 바로 뒤에 다시 넣고 원래 장은 지움(서식 · 그라데이션 그대로). 「되돌리기」는 원래 장을 다시 넣음
  const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  let items = null, itemKind = "", lastItem = null;

  async function loadItems() {
    if (items) return;
    items = (await (await fetch("items.json", { credentials: "same-origin" })).json()).filter((k) => k.items.length);
    $("#item-n").textContent = `· ${items.length}가지 · ${items.reduce((n, k) => n + k.items.length, 0)}개`;
    itemKind = items[0] ? items[0].code : "";
    const wrap = $("#item-kinds");
    wrap.replaceChildren(...items.map((k) => {
      const c = document.createElement("button");
      c.type = "button";
      c.className = "chip";
      c.title = k.desc;
      c.setAttribute("aria-pressed", String(k.code === itemKind));
      c.innerHTML = `${H().esc(k.name)}<small>${k.items.length}</small>`;
      c.onclick = () => {
        itemKind = k.code;
        wrap.querySelectorAll(".chip").forEach((x) => x.setAttribute("aria-pressed", String(x === c)));
        renderItems();
      };
      return c;
    }));
    renderItems();
  }
  function renderItems() {
    const k = items.find((x) => x.code === itemKind);
    if (!k) return;
    $("#items").replaceChildren(...k.items.map((it) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "item";
      b.title = `${k.name} — 지난 제안서 ${it.uses}곳에서 씀`;
      b.innerHTML = `<img loading="lazy" alt="" src="items/${it.id}.png"><span>${it.uses}곳에서 씀</span>`;
      b.onclick = () => insertItem(it, b);
      return b;
    }));
  }

  async function insertItem(it, btn) {
    if (!H().inOffice()) { H().toast("PowerPoint 안에서 끼워 넣을 수 있음", true); return; }
    if (btn) btn.disabled = true;
    try {
      let slideId = null, b64 = null, pos = null, tags = {};
      await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        if (!sel.items.length) return;
        const slide = sel.items[0];
        slideId = slide.id;
        slide.tags.load("items/key,items/value");
        const r = slide.exportAsBase64();
        await ctx.sync();
        b64 = r.value;
        tags = Object.fromEntries(slide.tags.items.map((t) => [t.key, t.value]));
        try {  // 고른 도형이 있으면 그 오른쪽에
          const shp = ctx.presentation.getSelectedShapes();
          shp.load("items/left,items/top,items/width,items/height");
          await ctx.sync();
          if (shp.items.length) { const s = shp.items[shp.items.length - 1]; pos = { x: s.left + s.width + 8, y: s.top }; }
        } catch (_) { /* 고른 도형 못 읽음 → 가운데 */ }
      });
      if (!b64) { H().toast("왼쪽에서 장을 하나 고르세요", true); return; }
      const zip = await JSZip.loadAsync(b64, { base64: true });
      const xmlName = Object.keys(zip.files).find((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
      const xml = await zip.file(xmlName).async("string");
      const pres = await zip.file("ppt/presentation.xml").async("string");
      const sw = +((/<p:sldSz[^>]*\bcx="(\d+)"/.exec(pres) || [])[1] || 12192000);
      const sh = +((/<p:sldSz[^>]*\bcy="(\d+)"/.exec(pres) || [])[1] || 6858000);
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const tree = doc.getElementsByTagNameNS(NS_P, "spTree")[0];
      if (!tree) throw new Error("장 구조를 못 읽음");
      let maxId = 0;
      for (const el of Array.from(doc.getElementsByTagName("*"))) if (el.localName === "cNvPr") maxId = Math.max(maxId, +el.getAttribute("id") || 0);
      const frag = new DOMParser().parseFromString(it.xml, "application/xml").documentElement;
      const node = doc.importNode(frag, true);
      for (const el of [node, ...Array.from(node.getElementsByTagName("*"))]) if (el.localName === "cNvPr") el.setAttribute("id", String(++maxId));
      const box = Array.from(node.children).find((c) => c.localName === "spPr" || c.localName === "grpSpPr");
      const xfrm = box && Array.from(box.children).find((c) => c.localName === "xfrm");
      const off = xfrm && Array.from(xfrm.children).find((c) => c.localName === "off");
      if (off) {
        let x = pos ? Math.round(pos.x * EMU) : Math.round((sw - it.w) / 2);
        let y = pos ? Math.round(pos.y * EMU) : Math.round((sh - it.h) / 2);
        x = Math.max(0, Math.min(x, sw - it.w));
        y = Math.max(0, Math.min(y, sh - it.h));
        off.setAttribute("x", String(x));
        off.setAttribute("y", String(y));
      }
      tree.appendChild(node);
      let out = new XMLSerializer().serializeToString(doc);
      if (!out.startsWith("<?xml")) out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + out;
      zip.file(xmlName, out);
      const nb64 = await zip.generateAsync({ type: "base64", compression: "DEFLATE" });
      const { at, newId } = await insertB64(nb64, "KeepSourceFormatting", { ...tags, HCG_ITEM: it.id, HCG_EDITED: today() }, slideId);
      if (!newId) throw new Error("다시 넣지 못함");
      await PowerPoint.run(async (ctx) => { ctx.presentation.slides.getItem(slideId).delete(); await ctx.sync(); });
      lastItem = { origB64: b64, newId, tags };
      $("#item-undo").hidden = false;
      H().toast(`${at}번 장에 Item 을 끼워 넣음 — ${pos ? "고른 도형 오른쪽" : "장 가운데"} · 글 자리는 ///`, false, { label: "되돌리기", run: undoItem });
    } catch (e) {
      console.error(e);
      H().toast("끼워 넣지 못함 — " + (e.message || e), true);
    } finally { if (btn) btn.disabled = false; }
  }
  async function undoItem() {
    if (!lastItem) return;
    const { origB64, newId, tags } = lastItem;
    try {
      const { newId: back } = await insertB64(origB64, "KeepSourceFormatting", tags, newId);
      if (back) await PowerPoint.run(async (ctx) => { ctx.presentation.slides.getItem(newId).delete(); await ctx.sync(); });
      lastItem = null;
      $("#item-undo").hidden = true;
      H().toast("Item 끼워 넣기 전으로 되돌림");
    } catch (e) {
      H().toast("되돌리지 못함 — " + (e.message || e), true);
    }
  }

  // ================= Component =================
  let parts = null;
  async function loadParts() {
    if (parts) return;
    const r = await fetch("parts.json", { credentials: "same-origin" });
    parts = await r.json();
    $("#parts-n").textContent = `Component ${parts.length}가지`;
    const grid = $("#parts");
    grid.replaceChildren(...parts.map((p) => {
      const b = document.createElement("button");
      b.className = "card"; b.type = "button";
      const v = p.variants[0];
      b.innerHTML = `<img loading="lazy" alt="" src="/api/thumb/${v.id}.jpg"><div class="t">${H().esc(p.name)}</div>
        <div class="m">Layout ${p.variants.length}가지 · 이 구조 장표 ${p.count}장</div>`;
      b.onclick = () => openPart(p);
      return b;
    }));
  }
  function openPart(p) {
    $("#parts").hidden = true;
    $("#part-view").hidden = false;
    $("#part-title").textContent = p.name;
    $("#part-src").textContent = `구조 ${p.struct} · Layout ${p.variants.length}가지 — 글 자리는 ///, 그림은 회색 상자(도형 · 구조만 가져옴)`;
    $("#part-vars").replaceChildren(...p.variants.map((v) => {
      const d = document.createElement("div");
      d.className = "card";
      d.innerHTML = `<img alt="" src="/api/img/${v.id}.jpg"><div class="t">${H().esc(v.title || p.name)}</div>
        <div class="m">${v.origin === "템플릿" ? "HCG 템플릿" : "지난 제안서에서 뽑은 Layout" + (v.uses > 1 ? ` · ${v.uses}곳에서 씀` : "")}</div>`;
      const btn = document.createElement("button");
      btn.className = "primary"; btn.type = "button"; btn.textContent = "이 Layout 삽입";
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
      const nc = H().getClient();
      H().toast(`${at}번 장에 Component 「${p.name}」 삽입 — 단 구성 · 문장 길이는 Revision 탭에서`, false,
        H().askBtn(`/hcg-restructure ${at}번 장은 빈 Component 「${p.name}」임(글 자리 ///, 회색 「그림」 상자).${nc ? ` 「${nc}」 제안` : " 이번 제안"} 내용으로 /// 자리마다 글을 채우고, 그림 상자는 알맞은 아이콘 · 그림으로 바꿀 안을 내 줘. 도형 · 서식은 그대로, 칸마다 글자 수 한도 안으로.`));
    } catch (e) {
      if (e && e.message !== "login") H().toast("넣지 못함 — " + (e.message || e), true);
    } finally { btn.disabled = false; btn.textContent = "이 Layout 삽입"; }
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
      if (res.status === Office.AsyncResultStatus.Succeeded) H().toast(`Icon 「${i.name}」 삽입`);
      else H().toast("아이콘을 넣지 못함 — " + (res.error && res.error.message), true);
    });
  }

  // ================= 칸 =================
  const cur = { slideId: null, no: 0, zip: null, xmlName: null, xml: null, slots: [], cols: [], tables: [], tags: {} };
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
        sel.items[0].tags.load("items/key,items/value");
        const r = sel.items[0].exportAsBase64();
        await ctx.sync();
        b64 = r.value;
        cur.tags = Object.fromEntries(sel.items[0].tags.items.map((t) => [t.key, t.value]));
      });
      if (!b64) { $("#slot-title").textContent = "왼쪽에서 장을 하나 고르세요"; $("#slot-sum").textContent = ""; $("#rev-steps").replaceChildren(); return; }
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
      await renderRevision(ov.length);
    } catch (e) {
      console.error(e);
      $("#slot-title").textContent = "장을 읽지 못함 — " + (e.message || e);
    } finally {
      reading = false;
      if (again) { again = false; readSlide(); }
    }
  }

  // ================= Revision — 고칠 기준 6가지를 차례로, 단추마다 무엇이 바뀌는지 먼저 적음 =================
  let NAMES = null;
  async function clientNames() {
    if (NAMES) return NAMES;
    try { NAMES = await (await fetch("check/client_names.json", { credentials: "same-origin" })).json(); } catch (_) { NAMES = []; }
    return NAMES;
  }
  const NUM_RX = /(?:\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s?(?:%|％|명|개사|개|억원|억|만원|천원|원|배|주|개월)(?!\s?\(가정\))/g;
  const escH = (s) => H().esc(s);
  function countIn(text, w) {
    if (!w) return 0;
    let n = 0, i = text.indexOf(w);
    while (i >= 0) { n++; i = text.indexOf(w, i + w.length); }
    return n;
  }

  function step(no, title, pill, pillCls, what, acts, extra = "") {
    const d = document.createElement("div");
    d.className = "rstep";
    d.innerHTML = `<div class="hd"><span class="no">${no}</span><b>${escH(title)}</b><span class="pill ${pillCls}">${escH(pill)}</span></div>
      <div class="what">${what}</div><div class="acts"></div>${extra}`;
    const box = d.querySelector(".acts");
    for (const [label, fn, ghost, disabled] of acts) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = ghost ? "ghost sm" : "primary sm";
      b.textContent = label;
      if (disabled) b.disabled = true;
      b.onclick = fn;
      box.appendChild(b);
    }
    if (!acts.length) box.remove();
    return d;
  }

  async function renderRevision(nOverlap) {
    const texts = cur.slots.filter((s) => s.chars > 0);
    const all = texts.map((s) => s.text || "").join("\n");
    const over = texts.filter((s) => s.overflow || s.wchars > s.limitHi);
    const src = cur.tags.HCG_SRC || "";
    $("#slot-title").textContent = `${cur.no}번 장`;
    $("#slot-sum").textContent = src ? ` · Library에서 삽입(${cur.tags.HCG_DOC || src})` : cur.tags.HCG_PART ? " · Component" : " · 직접 만든 장";
    const nc = H().getClient();
    const steps = [];
    const ask = (msg) => () => H().askClaude(msg);

    // 1 이전 고객명
    const names = await clientNames();
    const tagOld = String(cur.tags.HCG_OLD || "").split("|").filter(Boolean);
    const cand = [...new Set([...tagOld, ...names])].filter((n) => n && n !== nc && n.length >= 2);
    let found = cand.map((n) => [n, countIn(all, n)]).filter((x) => x[1] > 0);
    found = found.filter(([n]) => !found.some(([o]) => o !== n && o.includes(n)));  // 긴 이름 안의 짧은 이름은 뺌
    const nOld = found.reduce((a, x) => a + x[1], 0);
    const picks = nOld ? `<div class="opts">${found.map(([n, k]) => `<label class="tog"><input type="checkbox" class="oldpick" value="${escH(n)}" checked> 「${escH(n)}」 ${k}곳</label>`).join("")}</div>` : "";
    steps.push(step(1, "이전 고객명", nOld ? `${nOld}곳 남음` : "없음", nOld ? "r" : "g",
      nOld
        ? `체크한 이름을 이 장 글 · 도형 글에서 「${escH(nc || "고객사명")}」으로 바꿈. 사례 · 실적으로 든 회사 이름은 체크를 끄세요. 서식은 그대로, 표 칸 · 차트 글은 빠짐`
        : "이 장 글에서 이전 고객명을 찾지 못함. 그림 속 로고는 눈으로 확인",
      nOld ? [[nc ? `체크한 이름 → ${nc}` : "Library 탭 고객사명을 먼저 적으세요", () => {
        const sel = [...document.querySelectorAll("#rev-steps .oldpick")].filter((x) => x.checked).map((x) => [x.value, nc]);
        if (!sel.length) { H().toast("바꿀 이름을 체크하세요", true); return; }
        editAndReport(replacer(sel), "이전 고객명 교체");
      }, false, !nc]] : [], picks));

    // 2 수치 · 기간
    const nums = [...all.matchAll(NUM_RX)].map((m) => m[0]);
    const uniq = [...new Set(nums)];
    steps.push(step(2, "수치 · 기간", nums.length ? `${nums.length}곳` : "없음", nums.length ? "a" : "g",
      nums.length
        ? `${escH(uniq.slice(0, 6).join(" · "))}${uniq.length > 6 ? " …" : ""} — 지난 고객 값임. 「(가정) 표시」를 누르면 숫자 바로 뒤에 「(가정)」이 붙음(서식 그대로)`
        : "단위가 붙은 숫자를 찾지 못함",
      nums.length ? [
        ["(가정) 표시", () => editAndReport((t) => [...t.matchAll(NUM_RX)].map((m) => [m.index, m[0].length, m[0] + "(가정)"]), "(가정) 표시"), true],
        ["Claude에게 ›", ask(`/hcg-restructure ${cur.no}번 장의 수치 · 기간(${uniq.slice(0, 8).join(", ")})은 지난 고객 값임.${nc ? ` 「${nc}」` : " 새 고객"} 기준으로 바꿀 값이나 「(가정)」 표시 안을 장표는 그대로 두고 제안해 줘.`), true],
      ] : []));

    // 3 고객 context
    steps.push(step(3, "고객 context", "확인", "a",
      "지난 고객만의 사례 · 조직 · 제도 이름 · 산업 표현은 규칙으로 못 찾음. 누르면 이 장에서 찾아 바꿀 문장을 내 달라는 요청이 복사됨",
      [["Claude에게 ›", ask(`/hcg-restructure ${cur.no}번 장은 ${cur.tags.HCG_DOC ? `지난 제안서(${cur.tags.HCG_DOC})에서 가져온` : "지난"} 장표임.${nc ? ` 「${nc}」` : " 새 고객"}에 맞지 않는 고객 context(사례 · 조직 · 제도 이름 · 산업 표현 · 기간)를 찾아 칸마다 바꿀 문장을 제안해 줘. 도형 · 서식은 그대로.`), true]]));

    // 4 단 구성
    const ncol = cur.cols.length;
    const tables = [...new Set(cur.slots.filter((s) => s.kind === "cell").map((s) => s.tableNo))];
    cur.tables = tables;
    const extra4 = `<div class="opts"><label class="tog"><input type="checkbox" id="keep-orig" checked> 원래 장 남기기</label><label class="tog"><input type="checkbox" id="show-nums"> 칸 번호 보기</label></div>${tables.length ? '<div class="slot-tools" id="slot-tables"></div>' : ""}`;
    steps.push(step(4, "단 구성", ncol >= 2 ? `${ncol}단` : "단 없음", "g",
      ncol >= 2
        ? `가로로 늘어선 같은 폭 칸이 ${ncol}개. 누르면 ${ncol - 1}단 또는 ${ncol + 1}단짜리 새 장을 바로 뒤에 만듦(「원래 장 남기기」를 끄면 원래 장은 지움). 단을 늘리면 칸이 좁아져 문장이 넘칠 수 있음`
        : "가로로 늘어선 같은 폭 칸을 찾지 못함(좌우 비교 · 도형 묶음은 바꾸지 않음)",
      ncol >= 2 ? [
        [`${ncol - 1}단으로`, () => applyOp("cols", () => HCGSlot.setColumns(cur.xml, ncol - 1, cur.opts), `${ncol}단 → ${ncol - 1}단`), true, ncol <= 2],
        [`${ncol + 1}단으로`, () => applyOp("cols", () => HCGSlot.setColumns(cur.xml, ncol + 1, cur.opts), `${ncol}단 → ${ncol + 1}단`), true, ncol >= 5],
      ] : [], extra4));

    // 5 문장 길이 수정
    const list5 = over.length ? `<ul class="rlist">${over.slice(0, 8).map((s) => `<li data-no="${s.no}"><b>${s.no}</b> ${Math.round(s.wchars)}자 / 한도 ${Math.round(s.limitHi)} · ${escH((s.text || "").slice(0, 22))}…</li>`).join("")}</ul>` : "";
    steps.push(step(5, "문장 길이 수정", over.length ? `${over.length}칸 넘침` : nOverlap ? `겹침 ${nOverlap}` : "맞음", over.length || nOverlap ? "r" : "g",
      over.length
        ? "칸 크기보다 문장이 김(글자 수는 한글 1 · 영문 소문자 0.55로 어림). 누르면 넘친 칸을 한도 안으로 줄여 달라는 요청이 복사됨. 칸을 누르면 그 도형을 고름"
        : "모든 칸이 글자 수 한도 안",
      over.length ? [["Claude에게 줄이기 ›", ask(`/hcg-restructure ${cur.no}번 장 문장 길이 수정: ${over.slice(0, 10).map((s) => `${s.no}번 칸(「${(s.text || "").slice(0, 14)}…」) ${Math.round(s.limitLo)}~${Math.round(s.limitHi)}자`).join(", ")} 안으로 줄여 줘(폭 가중 글자 수). 도형 · 서식 그대로.`), true]] : [], list5));

    // 6 Layout 교체
    steps.push(step(6, "Layout 교체", src ? "같은 유형" : "Library 장만", src ? "g" : "a",
      src ? "같은 유형의 다른 Layout. 누르면 이 장 바로 뒤에 삽입(고객사명으로 바꿔서) — 원래 장과 비교한 뒤 하나를 지움" : "Library에서 삽입한 장만 유형을 알 수 있음",
      [], src ? '<div class="alts" id="rev-alts">불러오는 중…</div>' : ""));

    $("#rev-steps").replaceChildren(...steps);
    const sn = $("#show-nums");
    if (sn) sn.onchange = (e) => setNums(e.target.checked).catch((er) => H().toast("번호를 못 붙임 — " + er.message, true));
    if (tables.length) fillTables();
    $("#rev-steps").querySelectorAll(".rlist li").forEach((li) => {
      li.onclick = () => { const s = cur.slots.find((x) => x.no === +li.dataset.no); if (s) selectShape(s); };
    });
    if (src) loadAlts(src).catch((e) => { const a = $("#rev-alts"); if (a) a.textContent = "불러오지 못함 — " + (e.message || e); });
  }

  function fillTables() {
    const tw = $("#slot-tables");
    if (!tw) return;
    const tables = cur.tables;
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

  async function loadAlts(src) {
    const m = await H().api("meta/" + src);
    const { items } = await H().api("fam/" + m.fam);
    const alts = items.filter((x) => x.var !== m.var && x.id !== m.id).slice(0, 3);
    const box = $("#rev-alts");
    if (!box) return;
    if (!alts.length) { box.textContent = "같은 유형에 다른 Layout 없음"; return; }
    box.replaceChildren(...alts.map((x) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "alt";
      b.innerHTML = `<img alt="" src="/api/thumb/${x.id}.jpg"><span>${escH(x.name || x.fam_name || "")}</span>`;
      b.onclick = async () => {
        b.disabled = true;
        try {
          const r = await H().insertById(x.id, cur.slideId);
          H().toast(`${r.at}번 장에 다른 Layout 삽입 · 교체 ${r.n}곳 — 원래 장과 비교한 뒤 하나를 지우세요`);
        } catch (e) {
          H().toast("삽입하지 못함 — " + (e.message || e), true);
        } finally { b.disabled = false; }
      };
      return b;
    }));
  }

  // 글 고치기(서식 그대로) — 고른 장의 글 상자 · 묶음 안 도형. 표 · 그림 · 선은 건너뜀
  function replacer(pairs) {
    return (text) => {
      const out = [];
      for (const [from, to] of pairs) {
        if (!from || from === to) continue;
        let at = text.indexOf(from);
        while (at >= 0) { out.push([at, from.length, to]); at = text.indexOf(from, at + from.length); }
      }
      return out;
    };
  }
  async function editSlideText(fn) {
    let n = 0;
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(cur.slideId);
      const walk = async (shapes) => {
        shapes.load("items/type,items/name");
        await ctx.sync();
        const frames = [];
        for (const s of shapes.items) {
          if (s.type === "Group") { await walk(s.group.shapes); continue; }
          // 글 상자가 있는 도형만(차트 · SmartArt · 개체는 글 상자를 읽으면 오류가 나 장 전체 바꾸기가 실패함)
          if (!["GeometricShape", "TextBox", "Placeholder", "Freeform", "Callout"].includes(s.type)) continue;
          const tr = s.textFrame.textRange;
          tr.load("text");
          frames.push(tr);
        }
        await ctx.sync().catch(() => {});
        for (const tr of frames) {
          let text;
          try { text = tr.text; } catch (_) { continue; }
          if (!text) continue;
          const edits = fn(text).sort((a, b) => b[0] - a[0]);  // 뒤에서부터 바꿔 앞 위치가 안 밀리게
          for (const [at, len, to] of edits) { tr.getSubstring(at, len).text = to; n++; }
        }
        await ctx.sync();
      };
      await walk(slide.shapes);
    });
    return n;
  }
  async function editAndReport(fn, label) {
    try {
      const n = await editSlideText(fn);
      H().toast(`${label} ${n}곳 — Ctrl+Z 로 되돌릴 수 있음`);
      setTimeout(readSlide, 400);
    } catch (e) {
      console.error(e);
      H().toast(`${label} 못 함 — ` + (e.message || e), true);
    }
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
      const sn = $("#show-nums");
      if (sn && sn.checked) { sn.checked = false; await setNums(false); await readSlide(); }
      const keep = $("#keep-orig") ? $("#keep-orig").checked : true;
      const res = fn();
      if (!res || !res.xml || res.xml === cur.xml) throw new Error((res && res.warnings && res.warnings.join(" · ")) || "바꾸지 못함");
      cur.zip.file(cur.xmlName, res.xml);
      const b64 = await cur.zip.generateAsync({ type: "base64", compression: "DEFLATE" });
      const origId = cur.slideId;
      const { at, newId } = await insertB64(b64, "KeepSourceFormatting", { HCG_EDIT: kind, HCG_EDITED: today() }, origId);
      if (!keep && newId) {
        await PowerPoint.run(async (ctx) => { ctx.presentation.slides.getItem(origId).delete(); await ctx.sync(); });
      }
      let warn = (res.warnings || []).length ? ` · 확인: ${res.warnings.slice(0, 2).join(" · ")}` : "";
      if (/(\d)단 → (\d)단/.test(label) && +RegExp.$2 > +RegExp.$1) warn += " · 칸이 좁아져 문장이 넘칠 수 있음 — 「문장 길이 수정」을 확인하세요";
      H().toast(`${label} → ${keep ? at + "번 장에 새로" : "바꿈"}${warn}`);
      setTimeout(readSlide, 400);
    } catch (e) {
      console.error(e);
      H().toast("바꾸지 못함 — " + (e.message || e), true);
    }
  }

  function wire() {
    $("#part-back").onclick = () => { $("#part-view").hidden = true; $("#parts").hidden = false; };
    const seg = (item) => {
      $("#seg-comp").setAttribute("aria-pressed", String(!item));
      $("#seg-item").setAttribute("aria-pressed", String(item));
      $("#comp-pane").hidden = item;
      $("#item-pane").hidden = !item;
      if (item) { loadItems().catch((e) => H().toast("Item 을 불러오지 못함 — " + e.message, true)); loadIcons().catch(console.error); }
    };
    $("#seg-comp").onclick = () => seg(false);
    $("#seg-item").onclick = () => seg(true);
    $("#item-undo").onclick = undoItem;
    $("#iq").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(filterIcons, 200); });
    $("#icon-more").onclick = moreIcons;
    $("#slot-read").onclick = readSlide;
    document.addEventListener("hcg-tab", (e) => {
      if (e.detail === "parts") loadParts().catch(console.error);
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
