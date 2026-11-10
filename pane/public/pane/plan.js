/* 「HCG 장표」 Dummy 제안서 — Claude 가 만든 Dummy 장(장마다 유형 표시 HCG_PLAN_*)을 Library 장표로 채움.
   Claude 가 장을 직접 못 만들면 Claude 가 준 계획(HCG_PLAN JSON)을 붙여넣어 여기서 Dummy 장을 만듦. pane.js 의 window.HCGPane 을 씀 */
(() => {
  "use strict";
  const $ = (s) => document.querySelector(s);
  const H = () => window.HCGPane;
  const st = { dummies: [], picks: {}, alts: {} };

  // ---------- 파일 안 Dummy 장 읽기 ----------
  async function scan() {
    if (!H() || !H().inOffice()) return [];
    const out = [];
    await PowerPoint.run(async (ctx) => {
      const s = ctx.presentation.slides;
      s.load("items/id");
      await ctx.sync();
      for (const sl of s.items) sl.tags.load("items/key,items/value");
      await ctx.sync();
      s.items.forEach((sl, i) => {
        const t = Object.fromEntries(sl.tags.items.map((x) => [x.key, x.value]));
        if (t.HCG_PLAN_TYPE) out.push({ id: sl.id, at: i + 1, type: t.HCG_PLAN_TYPE, head: t.HCG_PLAN_HEAD || "", msg: t.HCG_PLAN_MSG || "", mod: t.HCG_PLAN_MOD || "" });
      });
    });
    st.dummies = out;
    return out;
  }
  async function banner() {
    try {
      const d = await scan();
      $("#plan-banner").hidden = !d.length;
      $("#plan-banner-n").textContent = d.length ? `이 파일에 Dummy 장 ${d.length}개` : "";
    } catch (_) { $("#plan-banner").hidden = true; }
  }

  // ---------- 유형마다 Library 후보 고르기(모듈 · 업종 맞는 것 → 많이 쓴 Layout) ----------
  const famCache = {};
  async function candidates(d, industry) {
    if (!famCache[d.type]) famCache[d.type] = (await H().api("fam/" + d.type)).items || [];
    const list = famCache[d.type];
    // 점수: Dummy 어깨끈 · 핵심 문장 낱말이 장표 이름에 겹치는 것 > 이 장 HR 모듈 > 프로젝트 주 모듈 > 업종 > 많이 쓴 Layout
    const words = toks(`${d.head} ${d.msg}`);
    const hasMod = (x, m) => !!m && (x.mods || []).some((y) => y.startsWith(m) || m.startsWith(y.slice(0, 2)));
    const score = (x) => {
      const nameT = toks(`${x.name || ""} ${x.head || ""}`);
      let kw = 0;
      for (const w of words) if (nameT.has(w)) kw++;
      return kw * 4 + (hasMod(x, d.mod) ? 2 : 0) + (hasMod(x, st.mainMod) ? 1.5 : 0) + (industry && x.industry === industry ? 0.5 : 0)
        + (x.tier === 1 ? 0.6 : 0) + (x.star ? 0.4 : 0) + 1 / (1 + (x.var_ord || 0));
    };
    return list.map((x) => [score(x), x]).sort((a, b) => b[0] - a[0]).map((p) => p[1]).slice(0, 8);
  }
  const STOP = new Set(["및", "등", "위한", "통해", "대한", "있음", "함", "수행", "방안", "제안", "프로젝트", "개요", "체계"]);
  function toks(s) {
    return new Set((String(s).toLowerCase().match(/[0-9a-z가-힣]{2,}/g) || []).filter((w) => !STOP.has(w)));
  }
  // 프로젝트 주 모듈 — Dummy 들의 HR 모듈 · 제도 설계 유형에서 가장 많은 것
  const TYPE_MOD = { T211: "평가", T212: "보상", T213: "직급", T214: "직무", T215: "조직", T216: "육성" };
  function mainMod(ds) {
    const c = {};
    for (const d of ds) { const m = d.mod || TYPE_MOD[d.type]; if (m) c[m] = (c[m] || 0) + 1; }
    return Object.keys(c).sort((a, b) => c[b] - c[a])[0] || "";
  }

  async function openPlan() {
    const pi = $("#plan-ind");  // 업종 목록은 Library 업종 거름에서 가져옴
    if (pi.options.length <= 1) for (const o of Array.from($("#fi").options).slice(1)) pi.add(new Option(o.textContent.replace(/\s*\(\d+\)$/, ""), o.value));
    pi.onchange = () => renderPlan().catch(console.error);
    $("#list-view").hidden = true;
    $("#plan-view").hidden = false;
    window.scrollTo(0, 0);
    await renderPlan();
  }
  async function renderPlan() {
    const d = await scan().catch(() => []);
    const box = $("#plan-list");
    $("#plan-sum").textContent = d.length ? `Dummy 장 ${d.length}개 — 유형마다 Library 장표를 골라 둠. 그림을 누르면 다음 후보로` : "이 파일에 Dummy 장이 없음 — 아래에 Claude 가 준 계획을 붙여넣어 만들 수 있음";
    $("#plan-fill").hidden = !d.length;
    box.replaceChildren();
    const ind = ($("#plan-ind") && $("#plan-ind").value) || "";
    st.mainMod = mainMod(d);
    const used = new Set();  // 같은 유형이 여러 장이어도 같은 장표를 두 번 넣지 않음
    const rows = d.map((x) => {
      const row = document.createElement("div");
      row.className = "prow";
      row.innerHTML = `<button type="button" class="pthumb" title="다음 후보"><span>불러오는 중</span></button>
        <div class="ptx"><b>${x.at} · ${H().esc(TYPE_NAME[x.type] || x.type)}</b><span>${H().esc(x.head)}</span><small>${H().esc(x.msg.slice(0, 60))}</small></div>`;
      box.appendChild(row);
      return row;
    });
    for (const [i, x] of d.entries()) {
      const row = rows[i];
      await candidates(x, ind).then((c0) => {
        const c = [...c0.filter((y) => !used.has(y.id)), ...c0.filter((y) => used.has(y.id))];
        if (c[0]) used.add(c[0].id);
        st.alts[x.id] = c;
        st.picks[x.id] = 0;
        const th = row.querySelector(".pthumb");
        const show = () => {
          const it = c[st.picks[x.id]];
          th.innerHTML = it ? `<img alt="" src="/api/thumb/${it.id}.jpg"><i>${st.picks[x.id] + 1}/${c.length}</i>` : "<span>맞는 장표 없음 → Component</span>";
        };
        th.onclick = () => { if (c.length) { st.picks[x.id] = (st.picks[x.id] + 1) % c.length; show(); } };
        show();
      }).catch(() => { row.querySelector(".pthumb").innerHTML = "<span>불러오지 못함</span>"; });
    }
  }

  async function fillAll() {
    const btn = $("#plan-fill");
    btn.disabled = true;
    const keep = $("#plan-keep").checked;
    let n = 0, skip = 0;
    try {
      // 뒤에서부터 채워 앞 장 번호가 밀리지 않게
      for (const d of [...st.dummies].reverse()) {
        const c = st.alts[d.id] || [];
        const it = c[st.picks[d.id] || 0];
        if (!it) { skip++; continue; }
        btn.textContent = `채우는 중… ${n + skip + 1} / ${st.dummies.length}`;
        await H().insertById(it.id, d.id);
        if (!keep) await PowerPoint.run(async (ctx) => { ctx.presentation.slides.getItem(d.id).delete(); await ctx.sync(); });
        n++;
      }
      const nc = H().getClient();
      H().toast(`Library 장표 ${n}장으로 채움${skip ? ` · 맞는 장표 없음 ${skip}장(Dummy 그대로)` : ""}`, false,
        H().askBtn(`/hcg-restructure 이 파일은 Dummy 계획대로 Library 장표를 채운 것임.${nc ? ` 고객 「${nc}」` : ""} 회의 메모 · 스콥에 맞게 장마다 핵심 문장과 /// · 지난 고객 내용을 고쳐 줘. 도형 · 서식은 그대로.`));
      setTimeout(renderPlan, 500);
      banner();
    } catch (e) {
      console.error(e);
      H().toast("채우지 못함 — " + (e.message || e), true);
    } finally {
      btn.disabled = false;
      btn.textContent = "모두 Library로 채우기";
    }
  }

  // ---------- 계획 붙여넣기 → Dummy 장 만들기(Claude 가 장을 직접 못 만들 때) ----------
  function parsePlan(text) {
    const m = text.match(/\{[\s\S]*\}/);
    if (!m) throw new Error("계획(JSON)을 찾지 못함 — Claude 답의 HCG_PLAN 블록 전체를 붙여넣으세요");
    const p = JSON.parse(m[0]);
    if (!Array.isArray(p.slides) || !p.slides.length) throw new Error("slides 목록이 없음");
    return p;
  }
  async function makeDummies() {
    let p;
    try { p = parsePlan($("#plan-paste").value); } catch (e) { H().toast(e.message, true); return; }
    const btn = $("#plan-make");
    btn.disabled = true;
    try {
      await PowerPoint.run(async (ctx) => {
        if (p.client) ctx.presentation.tags.add("HCG_CLIENT", String(p.client).slice(0, 250));
        for (const s of p.slides) {
          ctx.presentation.slides.add();
          await ctx.sync();
          const all = ctx.presentation.slides;
          all.load("items");
          await ctx.sync();
          const sl = all.items[all.items.length - 1];
          const head = `${s.ch || ""}    ${s.head || TYPE_NAME[s.type] || ""}`.trim();
          const t1 = sl.shapes.addTextBox(head, { left: 36, top: 18, width: 880, height: 26 });
          t1.textFrame.textRange.font.size = 14; t1.textFrame.textRange.font.color = "#C00000"; t1.textFrame.textRange.font.bold = true;
          const t2 = sl.shapes.addTextBox(s.msg || "///", { left: 36, top: 46, width: 880, height: 52 });
          t2.textFrame.textRange.font.size = 18; t2.textFrame.textRange.font.bold = true;
          const t3 = sl.shapes.addTextBox(`[Dummy] 유형: ${TYPE_NAME[s.type] || s.type}${s.note ? " · " + s.note : ""}\n///`, { left: 36, top: 120, width: 880, height: 300 });
          t3.textFrame.textRange.font.size = 14; t3.textFrame.textRange.font.color = "#7F7F7F";
          sl.tags.add("HCG_PLAN_TYPE", String(s.type || "T604"));
          sl.tags.add("HCG_PLAN_HEAD", String(s.head || "").slice(0, 250));
          sl.tags.add("HCG_PLAN_MSG", String(s.msg || "").slice(0, 250));
          if (s.mod) sl.tags.add("HCG_PLAN_MOD", String(s.mod).slice(0, 60));
          await ctx.sync();
        }
      });
      if (p.client && !H().getClient()) { $("#newClient").value = p.client; $("#newClient").dispatchEvent(new Event("change")); }
      H().toast(`Dummy 장 ${p.slides.length}개를 맨 뒤에 만듦 — 「모두 Library로 채우기」를 누르세요`);
      $("#plan-paste").value = "";
      await renderPlan();
      banner();
    } catch (e) {
      console.error(e);
      H().toast("Dummy 장을 만들지 못함 — " + (e.message || e), true);
    } finally { btn.disabled = false; }
  }

  let TYPE_NAME = {};
  async function loadTypes() {
    try {
      const t = await (await fetch("types.json", { credentials: "same-origin" })).json();
      TYPE_NAME = Object.fromEntries(t.types.map((x) => [x.code, x.name]));
    } catch (_) { TYPE_NAME = {}; }
  }

  function wire() {
    $("#plan-open").onclick = () => openPlan().catch(console.error);
    $("#plan-open2").onclick = () => openPlan().catch(console.error);
    $("#plan-back").onclick = () => { $("#plan-view").hidden = true; $("#list-view").hidden = false; };
    $("#plan-fill").onclick = fillAll;
    $("#plan-make").onclick = makeDummies;
    $("#plan-copy").onclick = () => H().askClaude("/hcg-dummy Dummy 제안서를 만들어 줘. 고객사: (이름 · 업종 · 규모)\n회의 내용: (메모 붙여넣기)\n주요 스콥: (예: 성과급 재설계, 직급체계 개편)\n기간: (예: 16주) · 분량: 일반 25장");
    loadTypes();
    const tryBanner = () => { if (H() && H().inOffice()) banner(); };
    setTimeout(tryBanner, 3000);
    document.addEventListener("hcg-tab", (e) => { if (e.detail === "store") tryBanner(); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wire); else wire();
})();
