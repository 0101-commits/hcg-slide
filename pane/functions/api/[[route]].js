// 「HCG 장표」 창고 API — Cloudflare Pages Functions
// 바인딩: DB(D1 검색 목록) · STORE(R2 한 장 PPTX · 그림)
// 환경 값: ACCESS_TEAM(예: hcg) · ACCESS_AUD(Access 앱 AUD) · ALLOWED_DOMAINS(예: e-hcg.com) · DEV_OPEN("1"이면 로그인 검사 생략 — 설정 전 시험용)

const COLS = "id,k,title,mods,mod,p,p_name,s,s_name,grade,star,kind,kind_name,action,client,industry,year,doc,page,bytes";
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
  if (best) where.push("(grade = '상' OR star = 1 OR kind <> '')");
  if (since) { where.push("year >= ?"); args.push(String(since)); }
  for (const term of q.split(/\s+/).filter(Boolean).slice(0, 6)) {
    where.push("(title LIKE ? OR text LIKE ? OR doc LIKE ? OR kind_name LIKE ? OR p_name LIKE ? OR s_name LIKE ? OR client LIKE ? OR industry LIKE ?)");
    const like = `%${term}%`;
    args.push(like, like, like, like, like, like, like, like);
  }
  const cond = where.length ? "WHERE " + where.join(" AND ") : "";
  const cnt = await env.DB.prepare(`SELECT COUNT(*) n FROM slides ${cond}`).bind(...args).first();
  const sql = `SELECT ${COLS} FROM slides ${cond}
    ORDER BY (kind <> '') DESC, CASE kind WHEN 'talenx' THEN 1 WHEN 'PSvsTI' THEN 2 WHEN 'PEST' THEN 3 WHEN '수행계획' THEN 4 WHEN '인력구조' THEN 5 WHEN '보상시뮬' THEN 6 ELSE 9 END, star DESC, (grade = '상') DESC, k LIMIT ? OFFSET ?`;
  args.push(limit, offset);
  const { results } = await env.DB.prepare(sql).bind(...args).all();
  return { total: cnt.n, items: results.map((r) => ({ ...r, mods: String(r.mods || "").split("|").filter(Boolean), star: !!r.star })) };
}

async function facets(env) {
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
    return json({ error: "로그인 확인 실패", detail: String(e) }, 500);
  }
  if (!me) return json({ error: "로그인이 필요함" }, 401);
  // 관리자 PC 로컬 시험(wrangler pages dev)에서만 — 장표 파일을 로컬 R2 에 빠르게 넣음
  if (request.method === "PUT" && route.startsWith("dev-put/") && env.DEV_OPEN === "1" && url.hostname === "localhost") {
    await env.STORE.put(route.slice(8), request.body, { httpMetadata: { contentType: request.headers.get("content-type") || "" } });
    return json({ ok: true });
  }
  if (request.method !== "GET") return json({ error: "GET 만 받음" }, 405);

  try {
    if (route === "me") return json({ email: me.email, dev: !!me.dev });
    if (route === "search") return json(await search(url, env));
    if (route === "facets") return json(await facets(env));
    let m = route.match(/^meta\/(s\d+)$/);
    if (m) {
      const r = await env.DB.prepare("SELECT * FROM slides WHERE id = ?").bind(m[1]).first();
      return r ? json({ ...r, mods: String(r.mods).split("|").filter(Boolean), star: !!r.star }) : json({ error: "없는 장표" }, 404);
    }
    m = route.match(/^slide\/(s\d+)\.pptx$/);
    if (m) return fromR2(env, `pptx/${m[1]}.pptx`, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    m = route.match(/^(img|thumb)\/(s\d+)\.jpg$/);
    if (m) return fromR2(env, `${m[1]}/${m[2]}.jpg`, "image/jpeg");
    return json({ error: "없는 주소" }, 404);
  } catch (e) {
    return json({ error: "창고 오류", detail: String(e) }, 500);
  }
}
