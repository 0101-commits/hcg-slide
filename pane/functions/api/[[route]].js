// 「HCG 장표」 창고 API — Cloudflare Pages Functions
// 바인딩: DB(D1 검색 목록) · STORE(R2 한 장 PPTX · 그림)
// 환경 값: ACCESS_TEAM(예: hcg) · ACCESS_AUD(Access 앱 AUD) · ALLOWED_DOMAINS(예: e-hcg.com) · DEV_OPEN("1"이면 로그인 검사 생략 — 설정 전 시험용)

const COLS = "id,k,title,name,head,mods,mod,p,p_name,s,s_name,grade,star,action,client,industry,year,doc,page,bytes,tier," +
  "fam,fam_name,fam_n,fam_v,fam_docs,var,var_n,var_ord,std,type,chap,tconf";
let jwksCache = null;

function json(data, status = 200, extra = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store", ...extra },
  });
}

function b64urlToBytes(s) {
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function getKeys(team) {
  if (jwksCache && jwksCache.team === team && Date.now() - jwksCache.at < 3600_000) return jwksCache.keys;
  const r = await fetch(`https://${team}.cloudflareaccess.com/cdn-cgi/access/certs`);
  if (!r.ok) throw new Error("certs " + r.status);
  const { keys } = await r.json();
  jwksCache = { team, at: Date.now(), keys };
  return keys;
}

// Access 가 붙여 주는 JWT 를 확인하고 메일 주소를 돌려줌. 실패하면 null.
async function whoami(request, env) {
  if (env.DEV_OPEN === "1") return { email: "dev@local", dev: true };
  const token =
    request.headers.get("cf-access-jwt-assertion") ||
    (request.headers.get("cookie") || "").match(/(?:^|;\s*)CF_Authorization=([^;]+)/)?.[1];
  if (!token || !env.ACCESS_TEAM || !env.ACCESS_AUD) return null;
  const [h, p, sig] = token.split(".");
  if (!sig) return null;
  const header = JSON.parse(new TextDecoder().decode(b64urlToBytes(h)));
  const payload = JSON.parse(new TextDecoder().decode(b64urlToBytes(p)));
  const keys = await getKeys(env.ACCESS_TEAM);
  const jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk) return null;
  const key = await crypto.subtle.importKey("jwk", jwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]);
  const ok = await crypto.subtle.verify("RSASSA-PKCS1-v1_5", key, b64urlToBytes(sig), new TextEncoder().encode(`${h}.${p}`));
  if (!ok) return null;
  const aud = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (!aud.includes(env.ACCESS_AUD)) return null;
  if (payload.exp && payload.exp * 1000 < Date.now()) return null;
  const email = String(payload.email || "").toLowerCase();
  const domains = String(env.ALLOWED_DOMAINS || "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  if (domains.length && !domains.some((d) => email.endsWith("@" + d))) return null;
  return { email };
}

async function search(url, env) {
  const q = (url.searchParams.get("q") || "").trim();
  const mod = url.searchParams.get("mod") || "";
  const p = url.searchParams.get("p") || "";
  const s = url.searchParams.get("s") || "";
  const ind = url.searchParams.get("ind") || "";
  const best = url.searchParams.get("best") === "1";
  const since = parseInt(url.searchParams.get("since") || "0", 10);
  const limit = Math.min(parseInt(url.searchParams.get("limit") || "60", 10) || 60, 200);
  const offset = Math.max(parseInt(url.searchParams.get("offset") || "0", 10) || 0, 0);
  const where = [];
  const args = [];
  if (mod) { where.push("mods LIKE ?"); args.push(`%|${mod}|%`); }
  if (p) { where.push("p = ?"); args.push(p); }
  if (s) { where.push("s = ?"); args.push(s); }
  if (ind) { where.push("industry = ?"); args.push(ind); }
  const chap = url.searchParams.get("chap") || "";
  if (chap) { where.push("chap = ?"); args.push(chap); }
  if (best) where.push("tier = 1");  // 재사용 「상」 · 별표 · 표준(색인으로 빨리)
  if (since) { where.push("year >= ?"); args.push(String(since)); }
  for (const term of q.split(/\s+/).filter(Boolean).slice(0, 6)) {
    where.push("(fam_name LIKE ? OR name LIKE ? OR head LIKE ? OR title LIKE ? OR text LIKE ? OR doc LIKE ? OR p_name LIKE ? OR s_name LIKE ? OR client LIKE ? OR industry LIKE ?)");
    const like = `%${term}%`;
    args.push(like, like, like, like, like, like, like, like, like, like);
  }
  const cond = where.length ? "WHERE " + where.join(" AND ") : "";
  // 찾는 말이 장표 이름(어깨끈)에 있으면 먼저
  const terms = q.split(/\s+/).filter(Boolean).slice(0, 6);
  const nameHit = terms.length ? terms.map(() => "(CASE WHEN fam_name LIKE ? OR name LIKE ? OR head LIKE ? THEN 1 ELSE 0 END)").join(" + ") : "0";
  const nameArgs = terms.flatMap((t) => [`%${t}%`, `%${t}%`, `%${t}%`]);
  // 묶음(같은 어깨끈 이름) 하나에 카드 하나 — 맞는 장 가운데 가장 많이 쓴 모양의 대표 장을 보여 줌
  const hit = `WITH hit AS (SELECT ${COLS}, ${nameHit} nh, ROW_NUMBER() OVER (PARTITION BY fam ORDER BY var_rep DESC, var_ord, tier, star DESC, year DESC, k) rn,
      COUNT(*) OVER (PARTITION BY fam) hit_n FROM slides ${cond})`;
  const hargs = [...nameArgs, ...args];
  const cnt = await env.DB.prepare(`SELECT COUNT(*) n, COUNT(DISTINCT fam) f, COUNT(DISTINCT CASE WHEN std = 1 THEN fam END) sf FROM slides ${cond}`).bind(...args).first();
  const fix = (r) => ({ ...r, mods: String(r.mods || "").split("|").filter(Boolean), star: !!r.star });
  // 자주 쓰는 표준 장표(여러 고객 제안서에 나온 묶음) — 첫 쪽에서만
  let std = [];
  if (!offset) {
    // 찾는 말이 없으면 유형을 제안서 흐름(장 → 유형 코드) 순으로, 있으면 이름 일치 → 맞는 장 수 순으로
    const order = terms.length ? "nh DESC, hit_n DESC, fam_docs DESC, k" : "CASE WHEN chap = '' OR chap IS NULL THEN 1 ELSE 0 END, chap, fam, fam_docs DESC, k";
    const { results } = await env.DB.prepare(`${hit} SELECT * FROM hit WHERE rn = 1 AND std = 1 ORDER BY ${order} LIMIT 80`).bind(...hargs).all();
    std = results.map(fix);
  }
  const { results } = await env.DB.prepare(`${hit} SELECT * FROM hit WHERE rn = 1 AND std = 0
    ORDER BY nh DESC, star DESC, (grade = '상') DESC, fam_docs DESC, fam_n DESC, k LIMIT ? OFFSET ?`).bind(...hargs, limit, offset).all();
  return { total: cnt.n, fams: cnt.f, std_fams: cnt.sf, rest_fams: cnt.f - cnt.sf, std, items: results.map(fix) };
}

// 묶음 하나의 모양들(모양마다 대표 장 하나) — 많이 쓴 모양 먼저
async function family(fam, env) {
  const { results } = await env.DB.prepare(`SELECT ${COLS} FROM slides WHERE fam = ? AND var_rep = 1 ORDER BY var_ord LIMIT 400`).bind(fam).all();
  return { items: results.map((r) => ({ ...r, mods: String(r.mods || "").split("|").filter(Boolean), star: !!r.star })) };
}

async function facets(env) {
  // 관리자가 올린 거름 개수(store/out/facets.json → R2 meta/facets.json)가 있으면 그것을 씀 — D1 을 덜 읽음
  const cached = await env.STORE.get("meta/facets.json");
  if (cached) return await cached.json();
  const [ps, ss, inds, total] = await Promise.all([
    env.DB.prepare("SELECT p, p_name, COUNT(*) n FROM slides GROUP BY p, p_name ORDER BY n DESC").all(),
    env.DB.prepare("SELECT s, s_name, COUNT(*) n FROM slides GROUP BY s, s_name ORDER BY n DESC").all(),
    env.DB.prepare("SELECT industry, COUNT(*) n FROM slides GROUP BY industry ORDER BY n DESC").all(),
    env.DB.prepare("SELECT COUNT(*) n FROM slides").first(),
  ]);
  const { results: modsRows } = await env.DB.prepare("SELECT mods FROM slides").all();
  const mods = {};
  for (const r of modsRows) for (const m of String(r.mods).split("|").filter(Boolean)) mods[m] = (mods[m] || 0) + 1;
  return { total: total.n, purposes: ps.results, structs: ss.results, industries: inds.results, mods };
}

async function fromR2(env, key, type) {
  const obj = await env.STORE.get(key);
  if (!obj) return json({ error: "없는 장표" }, 404);
  return new Response(obj.body, {
    headers: { "content-type": type, "cache-control": "private, max-age=86400", etag: obj.httpEtag },
  });
}

export async function onRequest({ request, env, params }) {
  const url = new URL(request.url);
  const route = (params.route || []).join("/");
  let me;
  try {
    me = await whoami(request, env);
  } catch (e) {
    console.error(e);
    return json({ error: "로그인 확인 실패" }, 500);
  }
  if (!me) return json({ error: "로그인이 필요함" }, 401);
  // 관리자 PC 로컬 시험(wrangler pages dev)에서만 — 장표 파일을 로컬 R2 에 빠르게 넣음
  if (request.method === "PUT" && route.startsWith("dev-put/") && env.DEV_OPEN === "1" && url.hostname === "localhost") {
    await env.STORE.put(route.slice(8), request.body, { httpMetadata: { contentType: request.headers.get("content-type") || "" } });
    return json({ ok: true });
  }
  // Claude Code 플러그인용 개인 열쇠 — 회사 메일로 로그인한 사람만 만들고 지움. 열쇠 원문은 만들 때 한 번만 돌려주고 해시만 둠
  if (request.method === "POST") {  // 다른 사이트에서 보낸 요청(CSRF) 막기
    const site = request.headers.get("sec-fetch-site");
    const origin = request.headers.get("origin");
    if ((site && site !== "same-origin") || (origin && origin !== url.origin)) return json({ error: "다른 사이트에서 온 요청" }, 403);
  }
  if (request.method === "POST" && route === "key/new") {
    const raw = new Uint8Array(32);
    crypto.getRandomValues(raw);
    const key = "hcgk_" + btoa(String.fromCharCode(...raw)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    const d = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(key));
    const hash = [...new Uint8Array(d)].map((x) => x.toString(16).padStart(2, "0")).join("");
    await env.DB.prepare("INSERT INTO keys (hash, email, created, last_used, revoked, tail) VALUES (?, ?, ?, NULL, 0, ?)")
      .bind(hash, me.email, Date.now(), key.slice(-4)).run();
    return json({ key, email: me.email });
  }
  if (request.method === "POST" && route === "key/revoke") {
    const tail = (url.searchParams.get("tail") || "").slice(0, 8);
    if (!tail) return json({ error: "지울 열쇠(tail)를 고르세요" }, 400);
    const r = await env.DB.prepare("UPDATE keys SET revoked = 1 WHERE email = ? AND tail = ?").bind(me.email, tail).run();
    return json({ ok: true, changed: r.meta?.changes ?? null });
  }
  if (request.method !== "GET") return json({ error: "GET 만 받음" }, 405);

  try {
    if (route === "me") return json({ email: me.email, dev: !!me.dev });
    if (route === "key/list") {
      const { results } = await env.DB.prepare("SELECT tail, created, last_used, revoked FROM keys WHERE email = ? ORDER BY created DESC LIMIT 20").bind(me.email).all();
      return json({ email: me.email, keys: results });
    }
    if (route === "search") return json(await search(url, env));
    if (route === "facets") return json(await facets(env));
    let m = route.match(/^meta\/(s\d+)$/);
    if (m) {
      const r = await env.DB.prepare("SELECT * FROM slides WHERE id = ?").bind(m[1]).first();
      return r ? json({ ...r, mods: String(r.mods).split("|").filter(Boolean), star: !!r.star }) : json({ error: "없는 장표" }, 404);
    }
    m = route.match(/^fam\/(f\d+|T\d{3})$/);  // f##### = 이름 묶음(유형 없는 장), T### = 유형
    if (m) return json(await family(m[1], env));
    // s = 창고 장표, p = 부품(글을 /// 로 지운 모양만)
    m = route.match(/^slide\/([sp]\d+)\.pptx$/);
    if (m) return fromR2(env, `pptx/${m[1]}.pptx`, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    m = route.match(/^(img|thumb)\/([sp]\d+)\.jpg$/);
    if (m) return fromR2(env, `${m[1]}/${m[2]}.jpg`, "image/jpeg");
    return json({ error: "없는 주소" }, 404);
  } catch (e) {
    console.error(e);
    return json({ error: "창고 오류" }, 500);
  }
}
