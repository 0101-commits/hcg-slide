// 「HCG 장표」 Claude Code 플러그인 전용 창고 통로 — 회사 메일 로그인(Access) 대신 사람마다 받은 개인 열쇠로 들어옴
// 열쇠는 /key 화면(회사 메일 로그인 뒤)에서 만들고, D1 keys 표에는 열쇠의 해시만 둠. 읽기만 됨(GET).
// Access: 이 경로(/kapi)만 「누구나 통과」 앱으로 열어 두고, 여기서 열쇠를 확인함(admin/kapi.mjs).

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });
}

async function sha256hex(s) {
  const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join("");
}

async function who(request, env) {
  const m = (request.headers.get("authorization") || "").match(/^Bearer\s+(hcgk_[A-Za-z0-9_-]{20,})$/);
  if (!m) return null;
  const h = await sha256hex(m[1]);
  const row = await env.DB.prepare("SELECT email, revoked, last_used FROM keys WHERE hash = ?").bind(h).first();
  if (!row || row.revoked) return null;
  if (!row.last_used || Date.now() - row.last_used > 3600_000) {
    await env.DB.prepare("UPDATE keys SET last_used = ? WHERE hash = ?").bind(Date.now(), h).run();
  }
  return { email: row.email };
}

async function fromR2(env, key, type) {
  const obj = await env.STORE.get(key);
  if (!obj) return json({ error: "없음", key }, 404);
  return new Response(obj.body, { headers: { "content-type": type, "cache-control": "private, max-age=3600", etag: obj.httpEtag } });
}

const PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export async function onRequest({ request, env, params }) {
  const route = (params.route || []).join("/");
  if (request.method !== "GET") return json({ error: "GET 만 받음" }, 405);
  let me;
  try {
    me = await who(request, env);
  } catch (e) {
    console.error(e);
    return json({ error: "열쇠 확인 실패" }, 500);
  }
  if (!me) return json({ error: "열쇠가 없거나 잘못됨 — /key 에서 새 열쇠를 받으세요" }, 401);
  try {
    if (route === "me") return json({ email: me.email });
    if (route === "version") {
      const heads = await Promise.all(["meta/kapi_meta.json", "meta/media_index.json", "meta/types.json", "frames/std.pptx"].map((k) => env.STORE.head(k)));
      return json({ meta: heads[0]?.httpEtag || null, media: heads[1]?.httpEtag || null, types: heads[2]?.httpEtag || null, frame: heads[3]?.httpEtag || null });
    }
    if (route === "meta") return fromR2(env, "meta/kapi_meta.json", "application/json; charset=utf-8");
    if (route === "media") return fromR2(env, "meta/media_index.json", "application/json; charset=utf-8");
    if (route === "types") return fromR2(env, "meta/types.json", "application/json; charset=utf-8");
    if (route === "frame/std.pptx") return fromR2(env, "frames/std.pptx", PPTX);
    const m = route.match(/^slide\/(s\d+)\.pptx$/);
    if (m) return fromR2(env, `pptx/${m[1]}.pptx`, PPTX);
    return json({ error: "없는 주소" }, 404);
  } catch (e) {
    console.error(e);
    return json({ error: "창고 오류" }, 500);
  }
}
