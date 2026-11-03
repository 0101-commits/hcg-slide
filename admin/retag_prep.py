# 2단계 다시 붙이기 재료 — 973장(「상」 ∪ 별표)을 문서가 쪼개지지 않게 4묶음으로 + 문서 목록
# 실행: ..\..\.venv\Scripts\python -X utf8 retag_prep.py   → work\retag\batch_1~4.tsv · docs.tsv
import collections, json, re

import store_build as B

W = B.WORK
R = W / "retag"
R.mkdir(exist_ok=True)


def clean(s):
    return " ".join(str(s or "").replace("\t", " ").split())


def base(name):
    return re.split(r"[\\/]", name)[-1]


def main(n=4):
    rows = [json.loads(l) for l in open(W / "ref" / "review" / "tags_all.jsonl", encoding="utf-8")]
    sl = {}
    for l in open(W / "ref" / "review" / "slides_all.jsonl", encoding="utf-8"):
        d = json.loads(l)
        sl[d["k"]] = d
    cand = [r for r in rows if r["rep"] == r["k"] and (r["re"] == "상" or r["star"])]
    cand.sort(key=lambda r: (r["doc"], r["i"]))
    size = -(-len(cand) // n)
    batches = [[] for _ in range(n)]
    b = 0
    for r in cand:
        if len(batches[b]) >= size and r["doc"] != batches[b][-1]["doc"] and b < n - 1:
            b += 1
        batches[b].append(r)
    for i, bt in enumerate(batches, 1):
        with open(R / f"batch_{i}.tsv", "w", encoding="utf-8") as f:
            f.write("sid\t문서\t목적\t구조\t제목\t글\n")
            for r in bt:
                s = sl.get(r["k"], {})
                cols = [f"s{r['k']:05d}", clean(base(r["name"]))[:80], B.PURPOSE.get(r["p"], r["p"]),
                        B.STRUCT.get(r["s"], r["s"]), clean(r.get("title"))[:100], clean(s.get("text"))[:600]]
                f.write("\t".join(cols) + "\n")
        print("batch", i, len(bt))
    cnt = collections.Counter(r["doc"] for r in cand)
    docs = collections.OrderedDict()
    for r in cand:
        docs.setdefault(r["doc"], r["name"])
    with open(R / "docs.tsv", "w", encoding="utf-8") as f:
        f.write("doc\t파일 경로(폴더/이름)\t장수\n")
        for d, nm in docs.items():
            f.write(f"{d}\t{clean(nm)}\t{cnt[d]}\n")
    print("docs", len(docs))


if __name__ == "__main__":
    main()
