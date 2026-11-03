// PowerPoint 작업창(WebView2)에 디버그 포트로 붙어 JS 를 실행 — 시험용
// 사용: node cdp_eval.mjs <port> "<식>"   (식은 async 함수 몸통처럼 await 가능, return 값 출력)
//       node cdp_eval.mjs <port> --list
//       node cdp_eval.mjs <port> --shot out.png
import { writeFileSync } from "node:fs";

const [port = "9229", expr = "--list", out] = process.argv.slice(2);
const targets = await (await fetch(`http://127.0.0.1:${port}/json`)).json();
if (expr === "--list") {
  for (const t of targets) console.log(t.type, t.url, t.title);
  process.exit(0);
}
const t = targets.find((x) => x.type === "page" && /\/pane\//.test(x.url)) || targets.find((x) => x.type === "page");
if (!t) { console.error("작업창 대상을 못 찾음"); process.exit(1); }
const ws = new WebSocket(t.webSocketDebuggerUrl);
let id = 0;
const pending = new Map();
ws.onmessage = (m) => {
  const d = JSON.parse(m.data);
  if (d.id && pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); }
};
await new Promise((r) => (ws.onopen = r));
const send = (method, params = {}) => new Promise((r) => { const i = ++id; pending.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
if (expr === "--shot") {
  const r = await send("Page.captureScreenshot", { format: "png" });
  writeFileSync(out || "pane.png", Buffer.from(r.result.data, "base64"));
  console.log("saved", out || "pane.png");
} else {
  const r = await send("Runtime.evaluate", { expression: `(async () => { ${expr} })()`, awaitPromise: true, returnByValue: true });
  console.log(JSON.stringify(r.result?.result?.value ?? r.result?.exceptionDetails ?? r, null, 1));
}
ws.close();
