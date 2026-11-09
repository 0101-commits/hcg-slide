/* 「HCG 장표」 4단계 — 점검 탭. checkops.js(window.HCGCheck)로 파일 전체를 규칙 점검하고, 누르면 그 장 · 그 칸으로 감 */
(() => {
  "use strict";
  const $ = (s) => document.querySelector(s);
  const H = () => window.HCGPane;
  const LV = { fix: "고칠 것", confirm: "확인할 것", ref: "참고" };
  const st = { res: null, slideIds: [], refs: null, running: false };

  async function refs() {
    if (st.refs) return st.refs;
    const [names, figures] = await Promise.all([
      fetch("check/client_names.json", { credentials: "same-origin" }).then((r) => (r.ok ? r.json() : [])),
      fetch("check/company_figures.json", { credentials: "same-origin" }).then((r) => (r.ok ? r.json() : null)),
    ]);
    st.refs = { names, figures };
    return st.refs;
  }

  // 지금 파일 전체를 PPTX 로 받음(조각으로)
  function getFile() {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
        const f = r.value, parts = [];
        let got = 0;
        const next = (i) => f.getSliceAsync(i, (s) => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { f.closeAsync(); return reject(s.error); }
          parts[i] = s.value.data;
          got++;
          if (got === f.sliceCount) {
            f.closeAsync();
            const total = parts.reduce((n, p) => n + p.length, 0);
            const out = new Uint8Array(total);
            let o = 0;
            for (const p of parts) { out.set(p, o); o += p.length; }
            resolve(out);
          } else next(i + 1);
        });
        next(0);
      });
    });
  }

  async function slideInfo() {
    let ids = [], tags = {};
    await PowerPoint.run(async (ctx) => {
      const s = ctx.presentation.slides;
      s.load("items/id");
      await ctx.sync();
      for (const sl of s.items) sl.tags.load("items/key,items/value");
      await ctx.sync();
      ids = s.items.map((x) => x.id);
      s.items.forEach((sl, i) => {
        const t = Object.fromEntries(sl.tags.items.map((x) => [x.key, x.value]));
        if (Object.keys(t).length) tags[i + 1] = t;
      });
    });
    return { ids, tags };
  }

  async function run() {
    if (!H().inOffice()) { H().toast("PowerPoint 안에서 점검할 수 있음", true); return; }
    if (!window.HCGCheck) { H().toast("점검 엔진을 불러오지 못함", true); return; }
    if (st.running) return;
    st.running = true;
    const btn = $("#chk-run");
    btn.disabled = true; btn.textContent = "점검 중…";
    try {
      const [bytes, info, rf] = await Promise.all([getFile(), slideInfo(), refs()]);
      st.slideIds = info.ids;
      const zip = await JSZip.loadAsync(bytes);
      const files = {};
      for (const name of Object.keys(zip.files)) {
        if (/\.(xml|rels)$/.test(name)) files[name] = await zip.file(name).async("string");
      }
      const opts = {
        client: H().getClient(),
        oldNames: rf.names,
        figures: rf.figures,
        publicRules: $("#chk-public").checked ? { 무기명: true } : null,
        slideTags: info.tags,
      };
      st.res = HCGCheck.run(files, opts);
      render();
    } catch (e) {
      console.error(e);
      H().toast("점검하지 못함 — " + (e.message || e), true);
    } finally {
      st.running = false;
      btn.disabled = false; btn.textContent = "다시 점검";
    }
  }

  function render() {
    const f = (st.res && st.res.findings) || [];
    const by = { fix: [], confirm: [], ref: [] };
    for (const x of f) (by[x.level] || by.ref).push(x);
    $("#chk-sum").innerHTML = ["fix", "confirm", "ref"].map((k) => `<span class="lv ${k}">${LV[k]} ${by[k].length}</span>`).join("");
    const wrap = $("#chk-list");
    wrap.replaceChildren();
    for (const k of ["fix", "confirm", "ref"]) {
      if (!by[k].length) continue;
      const h = document.createElement("h4");
      h.textContent = `${LV[k]} ${by[k].length}`;
      wrap.appendChild(h);
      for (const x of by[k]) wrap.appendChild(item(x));
    }
    if (!f.length) wrap.innerHTML = '<p class="empty">규칙으로 찾을 것이 없음. 문장 판단은 Claude(hcg-check)에게.</p>';
    // 같은 바꾸기(예: 이전 고객명 → 새 고객명)는 한 번에
    const pairs = new Map();
    const textFix = (x) => x.from && x.to && x.from !== x.to && !/레이아웃/.test(x.name || "");  // 레이아웃 이름은 글 바꾸기로 못 고침
    for (const x of f) if (textFix(x)) pairs.set(x.from + "\u0000" + x.to, [x.from, x.to]);
    const all = $("#chk-all");
    all.hidden = !pairs.size;
    all.textContent = `모두 바꾸기 (${pairs.size}가지)`;
    all.onclick = () => replaceAll([...pairs.values()]);
    // 규칙으로 못 고치는 것(고칠 것 중 글 바꾸기가 아닌 것 + 확인할 것)을 한 번에 Claude 에게
    const left = f.filter((x) => (x.level === "fix" && !textFix(x)) || x.level === "confirm");
    const ask = $("#chk-ask");
    ask.hidden = !left.length;
    ask.textContent = `남은 ${left.length}건 Claude에게 ›`;
    ask.onclick = () => {
      const lines = left.slice(0, 30).map((x) => `- ${x.slide ? x.slide + "번 장" : "전체"}${x.shape ? " · " + x.shape : ""} · ${x.name}: 「${(x.text || "").slice(0, 40)}」`);
      const nc = H().getClient();
      const msg = `/hcg-check 「HCG 장표」 점검에서 규칙으로 못 고친 ${left.length}건임${nc ? `(새 고객 「${nc}」)` : ""}. 문맥을 보고 장마다 고칠 안을 표로 내 줘. 고치는 것은 내가 고르면 할 것.\n` + lines.join("\n") + (left.length > 30 ? `\n… 외 ${left.length - 30}건` : "");
      H().askClaude(msg);
    };
  }

  function item(x) {
    const d = document.createElement("div");
    d.className = "fd " + x.level;
    d.innerHTML = `<div class="r1"><b>${H().esc(x.name || "항목 " + x.item)}</b><span class="where">${x.slide ? x.slide + "번 장" : "전체"}${x.shape ? " · " + H().esc(x.shape) : ""}</span></div>
      <div class="tx">${H().esc(x.text || "")}</div><div class="how">${H().esc(x.how || "")}</div><div class="bt"></div>`;
    const bt = d.querySelector(".bt");
    if (x.slide) {
      const go = document.createElement("button");
      go.className = "link"; go.type = "button"; go.textContent = "그 장으로 ›";
      go.onclick = () => goTo(x);
      bt.appendChild(go);
    }
    if (x.from && x.to && x.from !== x.to && !/레이아웃/.test(x.name || "")) {
      const rp = document.createElement("button");
      rp.className = "link"; rp.type = "button"; rp.textContent = `「${x.from}」→「${x.to}」 바꾸기`;
      rp.onclick = () => replaceAll([[x.from, x.to]], x.slide);
      bt.appendChild(rp);
    }
    if (x.level === "confirm" && /Claude/.test(x.how || "")) {
      const cl = document.createElement("button");
      cl.className = "link"; cl.type = "button"; cl.textContent = "Claude에게 ›";
      cl.onclick = async () => {
        const msg = `/hcg-check ${x.slide ? x.slide + "번 장 " : ""}「${(x.text || "").slice(0, 40)}」 — ${x.name}. 문맥을 보고 고칠 안을 내 줘.`;
        try { await navigator.clipboard.writeText(msg); H().toast("Claude 작업창에 붙여넣기(Ctrl+V) 하세요"); } catch (_) { H().toast(msg); }
      };
      bt.appendChild(cl);
    }
    return d;
  }

  async function goTo(x) {
    try {
      await PowerPoint.run(async (ctx) => {
        const id = st.slideIds[x.slide - 1];
        if (!id) return;
        ctx.presentation.setSelectedSlides([id]);
        await ctx.sync();
        if (!x.shape) return;
        const sh = ctx.presentation.slides.getItem(id).shapes;
        sh.load("items/id,items/name");
        await ctx.sync();
        const hit = sh.items.find((s) => s.name === x.shape);
        if (hit) { ctx.presentation.slides.getItem(id).setSelectedShapes([hit.id]); await ctx.sync(); }
      });
    } catch (e) { console.warn(e); }
  }

  // 글 바꾸기 — 글 조각(run) 서식은 그대로. 묶음 안 도형까지. 표 칸은 Claude 나 손으로.
  async function replaceAll(pairs, onlySlide) {
    let n = 0, skipped = 0;
    try {
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items/id");
        await ctx.sync();
        const targets = onlySlide ? [slides.items[onlySlide - 1]].filter(Boolean) : slides.items;
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
            for (const [from, to] of pairs) {
              let at = text.lastIndexOf(from);
              while (at >= 0) {
                tr.getSubstring(at, from.length).text = to;
                n++;
                text = text.slice(0, at) + to + text.slice(at + from.length);
                at = at > 0 ? text.lastIndexOf(from, at - 1) : -1;
              }
            }
          }
          await ctx.sync();
        };
        for (const sl of targets) {
          try { await walk(sl.shapes); } catch (e) { skipped++; console.warn(e); }
        }
      });
      H().toast(`바꾼 곳 ${n}${skipped ? ` · 못 읽은 장 ${skipped}` : ""} — 표 칸 · 차트 글은 다시 점검해 확인`);
      setTimeout(run, 600);
    } catch (e) {
      console.error(e);
      H().toast("바꾸지 못함 — " + (e.message || e), true);
    }
  }

  function wire() {
    $("#chk-run").onclick = run;
    document.addEventListener("hcg-tab", (e) => { if (e.detail === "check" && !st.res) run(); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wire); else wire();
})();
