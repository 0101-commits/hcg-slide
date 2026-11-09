# 창고 장마다 「다른 고객 그림」 수(opic)를 세어 store/out/opic.json 에 씀 — kapi.mjs data 가 kapi_meta 에 넣음
# 다른 고객 그림 = 한 고객 제안서(2건 이하)에만 나오는 그림. 작은 그림(8KB 아래, 아이콘 · 글머리표)은 셈에서 뺌.
# 플러그인은 이 수로 비슷한 장 가운데 그림이 적은 장을 먼저 고름(Draft 의 「다른 고객 그림」 딱지 줄이기).
# 실행: ..\.venv\Scripts\python admin\opic.py
import hashlib, json, os, re, sys, time, zipfile

APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
OUT = os.path.join(APP, "store", "out")
MIDX = os.path.join(APP, "..", "plan7", "run", "media_index.json")
IMG = re.compile(r"ppt/media/[^/]+\.(png|jpe?g|gif|bmp|emf|wmf|tiff?)$", re.I)


def main():
    t = time.time()
    midx = json.load(open(MIDX, encoding="utf-8"))
    meta = json.load(open(os.path.join(OUT, "meta.json"), encoding="utf-8"))
    res, miss = {}, 0
    for r in meta:
        p = os.path.join(OUT, "pptx", r["id"] + ".pptx")
        if not os.path.exists(p):
            miss += 1; continue
        n = 0
        try:
            with zipfile.ZipFile(p) as z:
                for zi in z.infolist():
                    if not IMG.match(zi.filename) or zi.file_size < 8000:
                        continue
                    e = midx.get(hashlib.md5(z.read(zi)).hexdigest())
                    cl = [x for x in (e or {}).get("clients", []) if x]
                    if e and len(cl) == 1 and e.get("ndocs", 9) <= 2:
                        n += 1
        except zipfile.BadZipFile:
            miss += 1; continue
        if n:
            res[r["id"]] = n
    json.dump(res, open(os.path.join(OUT, "opic.json"), "w", encoding="utf-8"))
    vals = list(res.values())
    print(json.dumps({"slides": len(meta), "with_pics": len(res), "miss": miss, "max": max(vals or [0]), "ge3": sum(v >= 3 for v in vals), "sec": round(time.time() - t, 1)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
