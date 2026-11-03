/* 「HCG 장표」 작업창 — 창고에서 찾기 · 바꿀 말 고치기 · 지금 파일에 그대로 넣기 */
(() => {
  "use strict";
  const $ = (s) => document.querySelector(s);
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const MOD_ORDER = ["평가·성과관리", "보상", "직급·승진", "직무", "조직 설계·진단", "공통", "리더십·육성", "조직문화", "인력·적정인력", "HR 시스템·솔루션", "채용", "노사·노무"];
  const MOD_SHORT = { "평가·성과관리": "평가", "직급·승진": "직급", "조직 설계·진단": "조직", "리더십·육성": "육성", "인력·적정인력": "인력", "HR 시스템·솔루션": "시스템", "노사·노무": "노사" };
  const state = { mod: "", inOffice: false, cur: null, timer: 0, seq: 0 };

  // ---------- 공통 ----------
  function toast(msg, isErr = false, action) {
    const t = $("#toast");
    t.className = "toast" + (isErr ? " err" : "");
    t.textContent = msg;
    if (action) {
      const b = document.createElement("button");
      b.textContent = action.label;
      b.onclick = action.run;
      t.appendChild(b);
    }
    t.hidden = false;
    clearTimeout(toast.h);
    toast.h = setTimeout(() => (t.hidden = true), isErr ? 9000 : 4500);
  }
  const relogin = { label: "다시 로그인", run: () => location.reload() };

  async function api(path, as = "json") {
    let r;
    try {
      r = await fetch("/api/" + path, { credentials: "same-origin", redirect: "manual" });
    } catch (e) {
      toast("창고에 닿지 못함 — 로그인이 끝났을 수 있음", true, relogin);
      throw e;
    }
    if (r.type === "opaqueredirect" || r.status === 401 || r.status === 403) {
      toast("로그인이 필요함", true, relogin);
      throw new Error("login");
    }
    if (!r.ok) {
      const msg = await r.text().catch(() => "");
      toast("창고 오류 " + r.status, true);
      throw new Error(r.status + " " + msg);
    }
    return as === "json" ? r.json() : r.arrayBuffer();
  }
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);

  // ---------- 새 고객명(파일마다 기억) ----------
  function getClient() { return $("#newClient").value.trim(); }
  function loadClient() {
    let v = "";
    try {
      if (state.inOffice) v = Office.context.document.settings.get("hcgNewClient") || "";
    } catch (_) { /* 설정 못 읽음 */ }
    if (!v) try { v = localStorage.getItem("hcgNewClient") || ""; } catch (_) { /* 저장소 없음 */ }
    $("#newClient").value = v;
  }
  function saveClient() {
    const v = getClient();
    try { localStorage.setItem("hcgNewClient", v); } catch (_) { /* 저장소 없음 */ }
    if (state.inOffice) {
      try {
        Office.context.document.settings.set("hcgNewClient", v);
        Office.context.document.settings.saveAsync();
      } catch (_) { /* 문서 설정 못 씀 */ }
    }
  }

  // ---------- 목록 ----------
  function card(it) {
    const b = document.createElement("button");
    b.className = "card";
    b.type = "button";
    const meta = it.kind_name ? `<span class="do">${esc(it.action)}</span>${it.client ? " · " + esc(it.client) : ""}` :
      `<span class="do">${esc(it.action)}</span> · ${esc(it.p_name || it.mod)}`;
    b.innerHTML = `<img loading="lazy" alt="" src="/api/thumb/${it.id}.jpg">
      <div class="t">${it.star ? '<span class="star">★</span> ' : ""}${esc(it.kind_name || it.title || "(제목 없음)")}</div>
      <div class="m">${meta}${it.year ? " · " + esc(it.year) : ""}${it.tier === 2 ? ' · <span class="tier2">재사용 중</span>' : ""}</div>`;
    b.title = `${it.title}\n${it.doc} ${it.page}쪽`;
    b.onclick = () => openInsert(it.id);
    return b;
  }

  async function runSearch(more = false) {
    const seq = ++state.seq;
    const p = new URLSearchParams();
    const PAGE = 60;
    state.offset = more ? (state.offset || 0) + PAGE : 0;
    const q = $("#q").value.trim();
    if (q) p.set("q", q);
    if (state.mod) p.set("mod", state.mod);
    if ($("#fp").value) p.set("p", $("#fp").value);
    if ($("#fs").value) p.set("s", $("#fs").value);
    if ($("#fi").value) p.set("ind", $("#fi").value);
    if ($("#fbest").checked) p.set("best", "1");
    if ($("#frecent").checked) p.set("since", String(new Date().getFullYear() - 3));
    p.set("limit", String(PAGE));
    p.set("offset", String(state.offset));
    const t0 = performance.now();
    const { items, total } = await api("search?" + p.toString());
    if (seq !== state.seq) return;
    state.lastMs = Math.round(performance.now() - t0);
    const std = items.filter((i) => i.kind);
    const rest = items.filter((i) => !i.kind);
    if (!more) {
      $("#std").replaceChildren(...std.map(card));
      $("#std-wrap").hidden = !std.length;
      $("#results").replaceChildren(...rest.map(card));
    } else {
      $("#std").append(...std.map(card));
      $("#results").append(...rest.map(card));
    }
    const shown = $("#std").children.length + $("#results").children.length;
    $("#res-title").textContent = `찾은 장표 ${total}장`;
    $("#more").hidden = shown >= total;
    $("#more").textContent = `더 보기 (${shown} / ${total})`;
    $("#empty").hidden = total > 0;
  }
  function searchSoon() {
    clearTimeout(state.timer);
    state.timer = setTimeout(() => runSearch().catch(() => {}), 220);
  }

  async function loadFacets() {
    const f = await api("facets");
    $("#total").textContent = `· 창고 ${f.total}장`;
    const wrap = $("#mods");
    const mk = (label, val, n) => {
      const c = document.createElement("button");
      c.className = "chip";
      c.type = "button";
      c.setAttribute("aria-pressed", String(state.mod === val));
      c.innerHTML = esc(label) + (n ? `<small>${n}</small>` : "");
      c.onclick = () => {
        state.mod = state.mod === val ? "" : val;
        wrap.querySelectorAll(".chip").forEach((x) => x.setAttribute("aria-pressed", "false"));
        c.setAttribute("aria-pressed", String(state.mod === val));
        if (!state.mod) wrap.firstChild.setAttribute("aria-pressed", "true");
        searchSoon();
      };
      return c;
    };
    const chips = [mk("전체", "", 0)];
    chips[0].setAttribute("aria-pressed", "true");
    for (const m of MOD_ORDER) if (f.mods[m]) chips.push(mk(MOD_SHORT[m] || m, m, f.mods[m]));
    wrap.replaceChildren(...chips);
    for (const r of f.purposes) $("#fp").add(new Option(`${r.p_name || r.p} (${r.n})`, r.p));
    for (const r of f.structs) $("#fs").add(new Option(`${r.s_name || r.s} (${r.n})`, r.s));
    for (const r of f.industries || []) if (r.industry) $("#fi").add(new Option(`${r.industry} (${r.n})`, r.industry));
  }

  // ---------- 넣기 화면 ----------
  function countIn(text, w) {
    if (!w) return 0;
    let n = 0, i = text.indexOf(w);
    while (i >= 0) { n++; i = text.indexOf(w, i + w.length); }
    return n;
  }
  function repRow(from = "", to = "", text = "") {
    const d = document.createElement("div");
    d.className = "rep";
    d.innerHTML = `<input class="from" value="${esc(from)}" placeholder="원래 말"><span class="arrow">→</span>
      <input class="to" value="${esc(to)}" placeholder="(그대로)"><button class="x" type="button" aria-label="빼기">✕</button>`;
    const n = countIn(text, from);
    if (from) {
      const c = document.createElement("div");
      c.className = "cnt";
      c.textContent = n ? `장표 글에서 ${n}곳` : "글 칸에서는 못 찾음(그림 · 묶음 이름일 수 있음)";
      d.appendChild(c);
    }
    d.querySelector(".x").onclick = () => d.remove();
    return d;
  }
  function numbersIn(text) {
    const rx = /(?:\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s?(?:%|％|명|개사|개|억원|억|만원|천원|원|배|주|개월)/g;
    const seen = new Set();
    for (const m of text.matchAll(rx)) {
      seen.add(m[0].replace(/\s+/g, ""));
      if (seen.size >= 6) break;
    }
    return [...seen];
  }

  async function openInsert(id) {
    const m = await api("meta/" + id);
    state.cur = m;
    $("#list-view").hidden = true;
    $("#insert-view").hidden = false;
    window.scrollTo(0, 0);
    $("#ins-title").textContent = (m.star ? "★ " : "") + (m.kind_name || m.title || "(제목 없음)");
    $("#ins-src").textContent = `원본 · ${m.doc} ${m.page}쪽${m.client ? " · " + m.client : ""}${m.industry ? " · " + m.industry : ""}${m.year ? " · " + m.year : ""}`;
    $("#ins-img").onerror = () => { $("#ins-img").onerror = null; $("#ins-img").src = `/api/thumb/${id}.jpg`; };
    $("#ins-img").src = `/api/img/${id}.jpg`;
    $("#ins-tags").innerHTML = [m.mod, m.p_name, m.s_name, m.grade ? `재사용 ${m.grade}` : "", m.action].filter(Boolean).map((t) => `<span>${esc(t)}</span>`).join("");
    const text = `${m.title} ${m.text} ${m.note || ""}`;
    const olds = String(m.olds || "").split("|").filter(Boolean);
    if (m.client && text.includes(m.client) && !olds.includes(m.client)) olds.unshift(m.client);
    const reps = [];
    const nc = getClient();
    // 원본 제안서의 고객만 새 고객명으로 미리 채움. 실적 · 투입 인력 장표의 다른 고객 이름은 실적 목록일 수 있어 그대로 둠
    const listy = ["P22", "P23"].includes(m.p);
    for (const o of olds) reps.push(repRow(o, o === m.client && !listy ? nc : "", text));
    const years = [...new Set(text.match(/(?<!\d)20[0-3]\d(?=\s?년|\.|\s|$)/g) || [])].sort().slice(-3);
    for (const y of years) reps.push(repRow(y, "", text));
    $("#reps").replaceChildren(...reps);
    $("#rep-count").textContent = reps.length ? `${reps.length}곳 후보` + (olds.length > 1 || listy ? " · 실적 목록의 고객 이름은 그대로 두세요" : "") : "· 찾은 후보 없음";
    const nums = numbersIn(text);
    $("#nums").innerHTML = nums.map((n) => `<li>${esc(n)} — 새 고객 값으로 바꾸거나 「(가정)」 표시</li>`).join("");
    $("#num-h").hidden = !nums.length;
    $("#do-insert").textContent = state.inOffice ? "넣기" : "PPTX 받기";
    await refreshHint();
  }
  function closeInsert() {
    $("#insert-view").hidden = true;
    $("#list-view").hidden = false;
    state.cur = null;
  }

  async function refreshHint() {
    if (!state.cur) return;
    if (!state.inOffice) {
      $("#ins-hint").textContent = "PowerPoint 밖에서 열려 있음 — 고친 장표를 파일로 받습니다.";
      return;
    }
    try {
      await PowerPoint.run(async (ctx) => {
        const all = ctx.presentation.slides;
        all.load("items/id");
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        if (!sel.items.length) {
          $("#ins-hint").textContent = `맨 뒤(${all.items.length + 1}번)에 넣음`;
        } else {
          const last = sel.items[sel.items.length - 1].id;
          const i = all.items.findIndex((s) => s.id === last) + 1;
          $("#ins-hint").textContent = `${i}번 장 뒤에 넣음 · 넣은 장에 출처 표시가 붙음`;
          $("#do-insert").textContent = `${i}번 장 뒤에 넣기`;
        }
      });
    } catch (e) {
      $("#ins-hint").textContent = "고른 장을 읽지 못함 — 맨 뒤에 넣음";
    }
  }

  // ---------- PPTX 글 바꾸기(서식 유지) ----------
  function replaceInXml(xml, pairs) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) return { xml, n: 0 };
    let n = 0;
    for (const p of Array.from(doc.getElementsByTagNameNS(A, "p"))) {
      const ts = Array.from(p.getElementsByTagNameNS(A, "t"));
      if (!ts.length) continue;
      for (const [from, to] of pairs) {
        let full = ts.map((t) => t.textContent).join("");
        let at = full.indexOf(from);
        while (at >= 0) {
          let pos = 0, first = true;
          const end = at + from.length;
          for (const t of ts) {
            const s = t.textContent, a = pos, b = pos + s.length;
            pos = b;
            const lo = Math.max(a, at), hi = Math.min(b, end);
            if (lo >= hi) continue;
            t.textContent = s.slice(0, lo - a) + (first ? to : "") + s.slice(hi - a);
            if (first && /[가-힣]/.test(to)) markKorean(t);
            first = false;
          }
          n++;
          full = ts.map((t) => t.textContent).join("");
          at = full.indexOf(from, at + to.length);
        }
      }
    }
    // 차트에 담긴 글 값(c:v)
    for (const v of Array.from(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/chart", "v"))) {
      let nv = v.textContent;
      for (const [from, to] of pairs) nv = nv.split(from).join(to);
      if (nv !== v.textContent) { v.textContent = nv; n++; }
    }
    // 도형 · 묶음 이름, 그림 설명
    for (const el of Array.from(doc.getElementsByTagName("*"))) {
      if (el.localName !== "cNvPr") continue;
      for (const att of ["name", "descr", "title"]) {
        const v = el.getAttribute(att);
        if (!v) continue;
        let nv = v;
        for (const [from, to] of pairs) nv = nv.split(from).join(to);
        if (nv !== v) { el.setAttribute(att, nv); n++; }
      }
    }
    let out = new XMLSerializer().serializeToString(doc);
    if (!out.startsWith("<?xml")) out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + out;
    return { xml: out, n };
  }
  function markKorean(t) {
    const r = t.parentNode;
    if (!r || r.localName !== "r") return;
    let rPr = Array.from(r.childNodes).find((c) => c.localName === "rPr");
    if (!rPr) {
      rPr = r.ownerDocument.createElementNS(A, "a:rPr");
      r.insertBefore(rPr, r.firstChild);
    }
    rPr.setAttribute("lang", "ko-KR");
    rPr.setAttribute("altLang", "en-US");
  }
  async function buildPptx(buf, pairs) {
    if (!pairs.length) return { buf, n: 0 };
    const zip = await JSZip.loadAsync(buf);
    let n = 0;
    const rx = /^ppt\/(slides\/slide\d+|notesSlides\/notesSlide\d+|slideMasters\/slideMaster\d+|slideLayouts\/slideLayout\d+|charts\/chart\d+|diagrams\/(data|drawing)\d+)\.xml$/;
    for (const name of Object.keys(zip.files)) {
      if (!rx.test(name)) continue;
      const r = replaceInXml(await zip.file(name).async("string"), pairs);
      if (r.n) { zip.file(name, r.xml); n += r.n; }
    }
    return { buf: await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" }), n };
  }
  function toBase64(buf) {
    const bytes = new Uint8Array(buf);
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  async function doInsert() {
    const m = state.cur;
    if (!m) return;
    const btn = $("#do-insert");
    btn.disabled = true;
    const label = btn.textContent;
    btn.textContent = "넣는 중…";
    try {
      const pairs = Array.from(document.querySelectorAll("#reps .rep"))
        .map((d) => [d.querySelector(".from").value.trim(), d.querySelector(".to").value.trim()])
        .filter(([f, t]) => f && t && f !== t);
      const raw = await api(`slide/${m.id}.pptx`, "buf");
      const { buf, n } = await buildPptx(raw, pairs);
      const fmt = document.querySelector('input[name="fmt"]:checked').value;
      if (!state.inOffice) {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
        a.download = `${m.id}_${(m.kind_name || m.title || "장표").slice(0, 30)}.pptx`;
        a.click();
        toast(`파일로 받음 · 바꾼 곳 ${n}`);
        return;
      }
      const b64 = toBase64(buf);
      let at = 0;
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items/id");
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        const before = new Set(slides.items.map((s) => s.id));
        const opts = { formatting: fmt };
        if (sel.items.length) opts.targetSlideId = sel.items[sel.items.length - 1].id;
        ctx.presentation.insertSlidesFromBase64(b64, opts);
        await ctx.sync();
        const after = ctx.presentation.slides;
        after.load("items/id");
        await ctx.sync();
        const added = after.items.filter((s) => !before.has(s.id));
        const today = new Date().toISOString().slice(0, 10);
        for (const s of added) {
          s.tags.add("HCG_SRC", m.id);
          s.tags.add("HCG_ADDED", today);
          s.tags.add("HCG_DOC", `${m.doc} ${m.page}`.slice(0, 250));
          const olds = String(m.olds || "").split("|").filter(Boolean);
          if (olds.length) s.tags.add("HCG_OLD", olds.join("|"));
          if (pairs.length) s.tags.add("HCG_REPLACED", pairs.map((p) => p.join("→")).join("|").slice(0, 250));
        }
        await ctx.sync();
        if (added.length) {
          at = after.items.findIndex((s) => s.id === added[0].id) + 1;
          ctx.presentation.setSelectedSlides([added[0].id]);
          await ctx.sync();
        }
      });
      toast(`${at}번 장으로 넣음 · 바꾼 곳 ${n} · 출처 표시 붙임`);
      closeInsert();
    } catch (e) {
      if (e && e.message !== "login") toast("넣지 못함 — " + (e.message || e), true);
      console.error(e);
    } finally {
      btn.disabled = false;
      btn.textContent = label;
    }
  }

  // ---------- 시작 ----------
  function wire() {
    document.querySelectorAll(".tabs button").forEach((b) => {
      b.onclick = () => {
        document.querySelectorAll(".tabs button").forEach((x) => x.setAttribute("aria-selected", String(x === b)));
        document.querySelectorAll(".tab").forEach((t) => (t.hidden = t.id !== "tab-" + b.dataset.tab));
        document.dispatchEvent(new CustomEvent("hcg-tab", { detail: b.dataset.tab }));
      };
    });
    $("#q").addEventListener("input", searchSoon);
    for (const s of ["#fp", "#fs", "#fi", "#fbest", "#frecent"]) $(s).addEventListener("change", searchSoon);
    $("#newClient").addEventListener("change", saveClient);
    $("#back").onclick = closeInsert;
    $("#cancel").onclick = closeInsert;
    $("#add-rep").onclick = () => $("#reps").appendChild(repRow("", "", ""));
    $("#do-insert").onclick = doInsert;
    $("#more").onclick = () => runSearch(true).catch(() => {});
  }

  async function start(inOffice) {
    state.inOffice = inOffice;
    // tools.js(부품 · 아이콘 · 칸 탭)가 같이 쓰는 것
    window.HCGPane = { api, toast, toBase64, esc, inOffice: () => state.inOffice, getClient };
    wire();
    loadClient();
    if (inOffice) {
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => refreshHint());
      } catch (_) { /* 선택 바뀜 알림 못 받음 */ }
    }
    try {
      const me = await api("me");
      $("#who").textContent = me.email;
      await loadFacets();
      await runSearch();
    } catch (e) {
      console.error(e);
    }
  }

  let started = false;
  const go = (v) => { if (!started) { started = true; start(v); } };
  if (window.Office && Office.onReady) {
    Office.onReady((info) => go(!!(info && info.host === Office.HostType.PowerPoint)));
    setTimeout(() => go(false), 4000);
  } else {
    go(false);
  }
})();
