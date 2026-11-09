# 유형 분류 준비 — 창고 장표마다 어깨끈 · 핵심 문장 · 본문 앞부분을 묶음 파일(TSV)로. 에이전트가 work/typing/out_NN.tsv 로 답함
import json, re
from pathlib import Path
APP = Path(__file__).resolve().parent.parent
W = APP.parent / "work" / "typing"
W.mkdir(parents=True, exist_ok=True)
meta = json.load(open(APP / "store" / "out" / "meta.json", encoding="utf-8"))
heads = {}
for l in open(APP / "store" / "out" / "heads.jsonl", encoding="utf-8"):
    d = json.loads(l); heads[d["id"]] = d
def clean(s): return re.sub(r"[\t\r\n]+", " ", s or "").strip()
rows = []
for m in sorted(meta, key=lambda x: x["k"]):
    h = heads.get(m["id"], {})
    body = clean((h.get("body") or [""])[0])[:120]
    text = clean(m.get("text", ""))
    for x in (m.get("head", ""), body):
        if x: text = text.replace(x, "")
    rows.append([m["id"], clean(m.get("head", ""))[:100], body, m.get("p_name", ""), m.get("s_name", ""), m.get("mod", ""), clean(text)[:200]])
N = 300
for i in range(0, len(rows), N):
    with open(W / f"batch_{i // N:02d}.tsv", "w", encoding="utf-8") as f:
        f.write("id\t어깨끈\t핵심문장\t목적\t구조\t모듈\t본문앞\n")
        for r in rows[i:i + N]: f.write("\t".join(r) + "\n")
t = json.load(open(APP / "admin" / "types.json", encoding="utf-8"))
ch = {c["id"]: c["name"] for c in t["chapters"]}
with open(W / "types.md", "w", encoding="utf-8") as f:
    f.write("# 유형 사전(52)\n\n| 코드 | 장 | 유형 | 무엇 | 예 |\n|---|---|---|---|---|\n")
    for x in t["types"]: f.write(f"| {x['code']} | {ch[x['ch']]} | {x['name']} | {x['desc']} | {x['ex']} |\n")
print(len(rows), "장 ·", (len(rows) + N - 1) // N, "묶음")
