# Component 마무리 — 후보(구조마다 24개)를 그림으로 확인한 구조(work/partcheck/out_*.tsv)로 다시 나누고 구조마다 20개로 줄임
# 실행: ..\..\.venv\Scripts\python -X utf8 parts_finalize.py --prep   → work/partcheck/list_NN.tsv (그림 확인 에이전트용)
#       ..\..\.venv\Scripts\python -X utf8 parts_finalize.py          → pane/public/pane/parts.json 다시 씀 · 안 쓰는 p#####.* 지움
import json
import sys
from pathlib import Path

APP = Path(__file__).resolve().parent.parent
OUT = APP / "store" / "out"
PJ = APP / "pane" / "public" / "pane" / "parts.json"
W = APP.parent / "work" / "partcheck"
KEEP = 20
sys.path.insert(0, str(Path(__file__).parent))


def prep(n=4):
    parts = json.load(open(PJ, encoding="utf-8"))
    rows = [(v["id"], p["code"]) for p in parts for v in p["variants"]]
    W.mkdir(parents=True, exist_ok=True)
    json.dump(parts, open(W / "pool.json", "w", encoding="utf-8"), ensure_ascii=False)
    per = (len(rows) + n - 1) // n
    for i in range(n):
        with open(W / f"list_{i:02d}.tsv", "w", encoding="utf-8") as f:
            for pid, code in rows[i * per:(i + 1) * per]:
                f.write(f"{pid}\t{code}\t{OUT / 'thumb' / (pid + '.jpg')}\n")
    print("후보", len(rows), "· 목록", n)


def finalize():
    from parts_build import PARTS
    import store_build as B
    pool = json.load(open(W / "pool.json", encoding="utf-8"))
    label = {}
    for f in sorted(W.glob("out_*.tsv")):
        for line in open(f, encoding="utf-8"):
            p = line.rstrip("\n").split("\t")
            if len(p) >= 2:
                label[p[0]] = p[1].strip()
    var = {v["id"]: dict(v, was=p["code"]) for p in pool for v in p["variants"]}
    cnt = {p["code"]: p["count"] for p in pool}
    by = {code: [] for code, _ in PARTS}
    moved = dropped = 0
    for pid, v in var.items():
        code = label.get(pid, v["was"])
        if code == "X" or code not in by:
            dropped += 1
            continue
        moved += code != v["was"]
        by[code].append(v)
    out = []
    keep = set()
    for code, name in PARTS:
        vs = sorted(by[code], key=lambda v: (v["origin"] != "템플릿", -v.get("uses", 0)))[:KEEP]
        for n, v in enumerate(vs, 1):
            v["title"] = f"{name} {n}"
            v.pop("was", None)
            keep.add(v["id"])
        if vs:
            out.append({"code": code, "name": name, "struct": B.STRUCT.get(code, code), "count": cnt.get(code, 0), "variants": vs})
        print(code, name, len(vs))
    PJ.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    for d, ext in (("pptx", "pptx"), ("img", "jpg"), ("thumb", "jpg")):
        for f in (OUT / d).glob(f"p*.{ext}"):
            if f.stem not in keep:
                f.unlink()
    print(f"parts.json {len(out)} 구조 · {len(keep)} Layout · 구조 바꿈 {moved} · 뺌 {dropped}")


if __name__ == "__main__":
    prep() if "--prep" in sys.argv else finalize()
