// Claude Code 플러그인 통로(/kapi) 준비 — 관리자 PC 에서
// 실행: node kapi.mjs [table|data|access|all]   (기본 all)
//   table  : D1 keys 표(개인 열쇠 해시) 만들기
//   data   : 플러그인이 받을 파일을 R2 에 올림 — meta/kapi_meta.json(검색용 줄인 목록) · meta/media_index.json · meta/types.json · frames/std.pptx(표지 · 간지 · 끝장 틀)
//   access : /kapi 경로만 「누구나 통과」(열쇠는 함수가 확인)
// 열쇠: hcg-proposal\.secrets\cf.env — deploy.mjs 와 같음
import { execFile } from "node:child_process";
import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const HERE = dirname(fileURLToPath(import.meta.url));
const APP = join(HERE, "..");
const PANE = join(APP, "pane");
const OUT = join(APP, "store", "out");
const RUN = join(APP, "..", "plan7", "run");
const TMP = join(APP, ".kapi");
const cfg = JSON.parse(readFileSync(join(APP, "config.json"), "utf8"));
const secret = Object.fromEntries(
  readFileSync(join(APP, "..", ".secrets", "cf.env"), "utf8").split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith("#") && l.includes("="))
    .map((l) => [l.slice(0, l.indexOf("=")).trim(), l.slice(l.indexOf("=") + 1).trim().replace(/^["']|["']$/g, "")]),
);
const env = { ...process.env, CLOUDFLARE_API_TOKEN: secret.CLOUDFLARE_API_TOKEN, CLOUDFLARE_ACCOUNT_ID: secret.CLOUDFLARE_ACCOUNT_ID, WRANGLER_SEND_METRICS: "false" };

async function wr(args) {
  const { stdout, stderr } = await run("npx", ["wrangler", ...args], { cwd: PANE, env, shell: true, maxBuffer: 64 << 20 });
  const out = (stdout + stderr).trim().split("\n").slice(-3).join("\n");
  console.log(out);
  return out;
}
async function cf(method, path, body) {
  const r = await fetch(`https://api.cloudflare.com/client/v4/accounts/${secret.CLOUDFLARE_ACCOUNT_ID}${path}`, {
    method, headers: { authorization: `Bearer ${secret.CLOUDFLARE_API_TOKEN}`, "content-type": "application/json" }, body: body ? JSON.stringify(body) : undefined,
  });
  const j = await r.json().catch(() => ({}));
  if (!j.success) throw new Error(`${method} ${path} → ${r.status} ${JSON.stringify(j.errors || j).slice(0, 400)}`);
  return j.result;
}

async function table() {
  await wr(["d1", "execute", cfg.d1_name, "--remote", "--command",
    `"CREATE TABLE IF NOT EXISTS keys (hash TEXT PRIMARY KEY, email TEXT NOT NULL, created INTEGER, last_used INTEGER, revoked INTEGER DEFAULT 0, tail TEXT); CREATE INDEX IF NOT EXISTS keys_email ON keys(email);"`]);
}

async function put(key, file, type) {
  await wr(["r2", "object", "put", `${cfg.r2_bucket}/${key}`, "--file", `"${file}"`, "--content-type", `"${type}"`, "--remote"]);
}

async function data() {
  mkdirSync(TMP, { recursive: true });
  const meta = JSON.parse(readFileSync(join(OUT, "meta.json"), "utf8"));
  const keep = ["id", "type", "client", "olds", "doc", "year", "industry", "name", "title", "std", "star", "grade", "tconf", "fam", "var", "mods", "page", "src"];
  const slim = meta.map((r) => ({ ...Object.fromEntries(keep.map((k) => [k, r[k]])), text: String(r.text || "").replace(/\s+/g, " ").slice(0, 600) }));
  const mp = join(TMP, "kapi_meta.json");
  writeFileSync(mp, JSON.stringify(slim));
  console.log("줄인 목록", slim.length, "장", (readFileSync(mp).length / 1e6).toFixed(1), "MB");
  // 표지 · 간지 Ⅰ~Ⅳ · 끝장 6장 틀(쿠리타한수 2026 제안서에서) — 합칠 때마다 고객 흔적은 Draft 단계가 지움
  const order = [1, 2, 6, 19, 50, 57].map((b, i) => ({ base: b, role: ["표지", "간지 Ⅰ", "간지 Ⅱ", "간지 Ⅲ", "간지 Ⅳ", "끝장"][i] }));
  const op = join(TMP, "frame_order.json");
  writeFileSync(op, JSON.stringify(order));
  const fp = join(TMP, "std.pptx");
  await run(join(APP, "..", ".venv", "Scripts", "python.exe"), ["-X", "utf8", join(RUN, "merge.py"), op, fp, join(RUN, "base_kurita.pptx")]);
  await put("meta/kapi_meta.json", mp, "application/json; charset=utf-8");
  await put("meta/media_index.json", join(RUN, "media_index.json"), "application/json; charset=utf-8");
  await put("meta/types.json", join(HERE, "types.json"), "application/json; charset=utf-8");
  await put("frames/std.pptx", fp, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
}

async function access() {
  const host = cfg.pages_url.replace(/^https?:\/\//, "");
  const d = `${host}/kapi`;
  const apps = await cf("GET", "/access/apps");
  const ex = apps.find((a) => a.domain === d);
  const body = { name: "HCG 장표 플러그인 통로 kapi", domain: d, type: "self_hosted", app_launcher_visible: false,
    policies: [{ name: "누구나(열쇠는 함수가 확인)", decision: "bypass", include: [{ everyone: {} }] }] };
  if (ex) await cf("PUT", `/access/apps/${ex.id}`, body); else await cf("POST", "/access/apps", body);
  console.log("Access 통과 경로", d);
}

const steps = { table, data, access, all: async () => { await table(); await data(); await access(); } };
const s = process.argv[2] || "all";
if (!steps[s]) { console.error("단계:", Object.keys(steps).join(" | ")); process.exit(1); }
await steps[s]();
