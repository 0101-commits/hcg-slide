// PowerPoint 작업창(WebView2)들 가운데 주소에 <match> 가 든 것에 붙어 조작 — 사용설명서 캡처 · 시험용
// 사용: node cdp_any.mjs <port> <match> --list
//       node cdp_any.mjs <port> <match> --eval "<식>"        (await 가능, return 값 출력)
//       node cdp_any.mjs <port> <match> --type "<글>"         (지금 커서 자리에 글 넣기)
//       node cdp_any.mjs <port> <match> --key Enter
import { readFileSync } from "node:fs";

const [port = "9339", match = "", cmd = "--list", arg = ""] = process.argv.slice(2);
const targets = await (await fetch(`http://127.0.0.1:${port}/json`)).json();
if (cmd === "--list") {
  for (const t of targets) console.log(t.type, "|", t.title, "|", t.url.slice(0, 120));
  process.exit(0);
}
const t = targets.find((x) => x.type === "page" && x.url.includes(match));
if (!t) { console.error("대상 없음:", match); process.exit(1); }
const ws = new WebSocket(t.webSocketDebuggerUrl);
let id = 0;
const pending = new Map();
ws.onmessage = (m) => { const d = JSON.parse(m.data); if (d.id && pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); } };
await new Promise((r) => (ws.onopen = r));
const send = (method, params = {}) => new Promise((res) => { const i = ++id; pending.set(i, res); ws.send(JSON.stringify({ id: i, method, params })); });
let out;
if (cmd === "--eval") {
  const src = arg.startsWith("@") ? readFileSync(arg.slice(1), "utf8") : arg;
  const r = await send("Runtime.evaluate", { expression: `(async () => { ${src} })()`, awaitPromise: true, returnByValue: true });
  out = r.result?.result?.value ?? r.result?.exceptionDetails?.exception?.description ?? JSON.stringify(r.result).slice(0, 500);
} else if (cmd === "--type") {
  const text = arg.startsWith("@") ? readFileSync(arg.slice(1), "utf8") : arg;
  await send("Input.insertText", { text });
  out = "typed " + text.length;
} else if (cmd === "--key") {
  const key = arg || "Enter";
  const code = key === "Enter" ? 13 : 0;
  await send("Input.dispatchKeyEvent", { type: "keyDown", key, code: key, windowsVirtualKeyCode: code, text: key === "Enter" ? "\r" : undefined });
  await send("Input.dispatchKeyEvent", { type: "keyUp", key, code: key, windowsVirtualKeyCode: code });
  out = "key " + key;
}
console.log(typeof out === "string" ? out : JSON.stringify(out));
ws.close();
process.exit(0);
