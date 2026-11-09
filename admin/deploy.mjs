// Cloudflare 에 창고 · 작업창 올리기 (관리자 PC 에서)
// 실행: node deploy.mjs [setup|data|pages|access|all]   (기본 all)
//   setup  : D1 · R2 · Pages 프로젝트 만들기(있으면 건너뜀)
//   data   : D1 검색 목록 · R2 장표 파일 올리기 (store/out)
//   pages  : 작업창 화면 + API 배포
//   access : 회사 메일 로그인(Zero Trust Access) 설정 → AUD 를 wrangler.toml 에 넣고 다시 배포
// 열쇠: hcg-proposal\.secrets\cf.env (CLOUDFLARE_API_TOKEN · CLOUDFLARE_ACCOUNT_ID) — 기존 wrangler 로그인은 건드리지 않음
import { execFile } from "node:child_process";
import { readFileSync, writeFileSync, existsSync, readdirSync, statSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const HERE = dirname(fileURLToPath(import.meta.url));
const APP = join(HERE, "..");
const PANE = join(APP, "pane");
const OUT = join(APP, "store", "out");
const SECRETS = join(APP, "..", ".secrets", "cf.env");
const STATE = join(APP, ".deploy.json");
const cfg = JSON.parse(readFileSync(join(APP, "config.json"), "utf8"));

if (!existsSync(SECRETS)) {
  console.error("열쇠 파일이 없음:", SECRETS);
  process.exit(1);
}
const secret = Object.fromEntries(
  readFileSync(SECRETS, "utf8").split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith("#") && l.includes("="))
    .map((l) => [l.slice(0, l.indexOf("=")).trim(), l.slice(l.indexOf("=") + 1).trim().replace(/^["']|["']$/g, "")]),
);
const TOKEN = secret.CLOUDFLARE_API_TOKEN;
const ACCOUNT = secret.CLOUDFLARE_ACCOUNT_ID;
if (!TOKEN || !ACCOUNT) {
  console.error("cf.env 에 CLOUDFLARE_API_TOKEN · CLOUDFLARE_ACCOUNT_ID 가 있어야 함");
  process.exit(1);
}
const env = { ...process.env, CLOUDFLARE_API_TOKEN: TOKEN, CLOUDFLARE_ACCOUNT_ID: ACCOUNT, WRANGLER_SEND_METRICS: "false" };
const state = existsSync(STATE) ? JSON.parse(readFileSync(STATE, "utf8")) : {};
const save = () => writeFileSync(STATE, JSON.stringify(state, null, 1));

async function wr(args, { quiet = false } = {}) {
  const { stdout, stderr } = await run("npx", ["wrangler", ...args], { cwd: PANE, env, shell: true, maxBuffer: 64 << 20 });
  if (!quiet) process.stdout.write(stdout.split("\n").slice(-6).join("\n") + "\n");
  return stdout + stderr;
}
async function cf(method, path, body) {
  const r = await fetch(`https://api.cloudflare.com/client/v4/accounts/${ACCOUNT}${path}`, {
    method,
    headers: { authorization: `Bearer ${TOKEN}`, "content-type": "application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  const j = await r.json().catch(() => ({}));
  if (!j.success) throw new Error(`${method} ${path} → ${r.status} ${JSON.stringify(j.errors || j).slice(0, 400)}`);
  return j.result;
}
function setToml(re, line) {
  const p = join(PANE, "wrangler.toml");
  const t = readFileSync(p, "utf8");
  if (!re.test(t)) throw new Error("wrangler.toml 에서 못 찾음: " + re);
  writeFileSync(p, t.replace(re, line));
}

async function setup() {
  // D1
  const dbs = await cf("GET", "/d1/database?per_page=100");
  let db = dbs.find((d) => d.name === cfg.d1_name);
  if (!db) db = await cf("POST", "/d1/database", { name: cfg.d1_name });
  state.database_id = db.uuid;
  setToml(/database_id = ".*"/, `database_id = "${db.uuid}"`);
  console.log("D1", cfg.d1_name, db.uuid);
  // R2
  const { buckets } = await cf("GET", "/r2/buckets");
  if (!buckets.some((b) => b.name === cfg.r2_bucket)) await cf("POST", "/r2/buckets", { name: cfg.r2_bucket });
  console.log("R2", cfg.r2_bucket);
  // Pages
  const projects = await cf("GET", "/pages/projects");
  let pr = projects.find((p) => p.name === cfg.pages_project);
  if (!pr) pr = await cf("POST", "/pages/projects", { name: cfg.pages_project, production_branch: "main" });
  state.pages_url = "https://" + pr.subdomain;
  console.log("Pages", state.pages_url);
  if (state.pages_url !== cfg.pages_url) {
    cfg.pages_url = state.pages_url;
    writeFileSync(join(APP, "config.json"), JSON.stringify(cfg, null, 2) + "\n");
    console.log("config.json pages_url 고침 →", cfg.pages_url);
  }
  save();
}

async function data() {
  if (process.env.SKIP_D1 !== "1") {  // SKIP_D1=1 이면 R2 만
    await wr(["d1", "execute", cfg.d1_name, "--remote", "--yes", "--file", "schema.sql"]);
    // 판이 바뀌며 더한 칸(2단계 업종) — 이미 있으면 건너뜀
    for (const m of ["migrate_2.sql", "migrate_3.sql", "migrate_4.sql", "migrate_5.sql"]) {
      if ((state.migrations || []).includes(m)) continue;
      try { await wr(["d1", "execute", cfg.d1_name, "--remote", "--yes", "--file", m], { quiet: true }); }
      catch (e) { if (!/duplicate column/i.test(String(e))) throw e; }
      state.migrations = [...(state.migrations || []), m];
      save();
    }
    await wr(["d1", "execute", cfg.d1_name, "--remote", "--yes", "--file", join(OUT, "d1_seed.sql")]);
  }
  // R2 — API 로 바로 올림(4개 동시, 한도에 걸리면 쉬었다 다시). 크기 · 고친 시각이 같으면 건너뜀
  state.r2 = state.r2 || {};
  const jobs = [];
  // 「중」(tier 2) 장표는 큰 미리보기(img)를 올리지 않음 — 작업창이 작은 그림(thumb)으로 대신 보여 줌(올리는 수 ⅓ 줄임)
  const tier = Object.fromEntries((existsSync(join(OUT, "meta.json")) ? JSON.parse(readFileSync(join(OUT, "meta.json"), "utf8")) : []).map((m) => [m.id, m.tier]));
  for (const id of readdirSync(join(OUT, "pptx")).filter((f) => f.endsWith(".pptx")).map((f) => f.slice(0, -5))) {
    // 창고(meta) · 부품(p#####) 이 아닌 파일 = 부품을 만들 때 쓴 템플릿 원본 → 올리지 않음
    if (!(id in tier) && !id.startsWith("p") && !existsSync(join(OUT, "thumb", `${id}.jpg`))) continue;
    jobs.push([`pptx/${id}.pptx`, join(OUT, "pptx", `${id}.pptx`), "application/vnd.openxmlformats-officedocument.presentationml.presentation"]);
    if (tier[id] !== 2 && existsSync(join(OUT, "img", `${id}.jpg`))) jobs.push([`img/${id}.jpg`, join(OUT, "img", `${id}.jpg`), "image/jpeg"]);
    if (existsSync(join(OUT, "thumb", `${id}.jpg`))) jobs.push([`thumb/${id}.jpg`, join(OUT, "thumb", `${id}.jpg`), "image/jpeg"]);
  }
  if (existsSync(join(OUT, "facets.json"))) jobs.unshift(["meta/facets.json", join(OUT, "facets.json"), "application/json"]);
  let done = 0, skip = 0, fail = 0;
  const base = `https://api.cloudflare.com/client/v4/accounts/${ACCOUNT}/r2/buckets/${cfg.r2_bucket}/objects/`;
  const pool = Array.from({ length: 4 }, async () => {
    while (jobs.length) {
      const [key, file, type] = jobs.shift();
      const st = statSync(file);
      const sig = `${st.size}-${Math.round(st.mtimeMs)}`;
      if (state.r2[key] === sig) { skip++; continue; }
      let ok = false;
      for (let t = 0; t < 8 && !ok; t++) {
        try {
          const r = await fetch(base + key, { method: "PUT", headers: { authorization: `Bearer ${TOKEN}`, "content-type": type }, body: readFileSync(file) });
          ok = r.ok;
          if (!ok) {
            const body = (await r.text()).slice(0, 120);
            // 429 = 너무 빠름(계정 API 한도 5분 1,200번) → 쉬었다가 다시
            await new Promise((res) => setTimeout(res, r.status === 429 ? 15000 * (t + 1) : 2000));
            if (r.status !== 429) console.error("다시", key, r.status, body);
          }
        } catch (e) { console.error("다시", key, String(e).slice(0, 120)); }
      }
      if (ok) { done++; state.r2[key] = sig; } else { fail++; console.error("실패", key); }
      if ((done + fail) % 200 === 0) { console.log(`R2 ${done + fail} 올림 · ${skip} 건너뜀`); save(); }
    }
  });
  await Promise.all(pool);
  save();
  console.log(`R2 올림 ${done} · 건너뜀 ${skip} · 실패 ${fail}`);
  if (fail) process.exitCode = 1;
}

async function pages() {
  setToml(/DEV_OPEN = ".*"/, `DEV_OPEN = "0"`);
  await wr(["pages", "deploy", "public", "--project-name", cfg.pages_project, "--branch", "main", "--commit-dirty=true"]);
}

async function access() {
  const org = await cf("GET", "/access/organizations");
  const team = org.auth_domain.replace(".cloudflareaccess.com", "");
  state.access_team = team;
  // 로그인 방법: 메일로 받는 숫자(One-time PIN)
  const idps = await cf("GET", "/access/identity_providers");
  let otp = idps.find((i) => i.type === "onetimepin");
  if (!otp) otp = await cf("POST", "/access/identity_providers", { name: "회사 메일 숫자", type: "onetimepin", config: {} });
  const host = new URL(cfg.pages_url).host;
  const apps = await cf("GET", "/access/apps");
  const include = cfg.allowed_email_domains.map((d) => ({ email_domain: { domain: d } }));
  // 1) 작업창 · API 전체 — 회사 메일만
  let app = apps.find((a) => a.domain === host);
  const appBody = {
    name: "HCG 장표", domain: host, type: "self_hosted", session_duration: "720h",
    allowed_idps: [otp.id], auto_redirect_to_identity: true, app_launcher_visible: false,
    policies: [{ name: "회사 메일", decision: "allow", include }],
  };
  app = app ? await cf("PUT", `/access/apps/${app.id}`, appBody) : await cf("POST", "/access/apps", appBody);
  // 2) 리본 아이콘 · 도움말 — PowerPoint 가 로그인 없이 받아야 함
  for (const path of ["icons", "help"]) {
    const d = `${host}/${path}`;
    const ex = apps.find((a) => a.domain === d);
    const body = { name: `HCG 장표 공개 ${path}`, domain: d, type: "self_hosted", app_launcher_visible: false,
      policies: [{ name: "누구나", decision: "bypass", include: [{ everyone: {} }] }] };
    if (ex) await cf("PUT", `/access/apps/${ex.id}`, body); else await cf("POST", "/access/apps", body);
  }
  state.access_aud = app.aud;
  save();
  setToml(/ACCESS_TEAM = ".*"/, `ACCESS_TEAM = "${team}"`);
  setToml(/ACCESS_AUD = ".*"/, `ACCESS_AUD = "${app.aud}"`);
  setToml(/ALLOWED_DOMAINS = ".*"/, `ALLOWED_DOMAINS = "${cfg.allowed_email_domains.join(",")}"`);
  cfg.access_team = team;
  writeFileSync(join(APP, "config.json"), JSON.stringify(cfg, null, 2) + "\n");
  console.log("Access", team, "AUD", app.aud.slice(0, 8) + "…");
  await pages();
}

const step = process.argv[2] || "all";
const steps = { setup, data, pages, access, all: async () => { await setup(); await data(); await pages(); await access(); } };
if (!steps[step]) { console.error("단계:", Object.keys(steps).join(" | ")); process.exit(1); }
steps[step]().catch((e) => { console.error("실패:", e.message); process.exit(1); });
