# 창고 장표 이름 · 묶음 — 어깨끈(제목 개체 틀 「대분류    소분류」)의 마지막 칸을 장표 이름으로 쓰고,
# 이름이 같은 장표를 한 묶음(fam)으로, 묶음 안에서 거의 같은 모양(그림 dHash · 글 겹침)을 한 모양(var)으로 묶음.
# 자주 쓰는 표준 장표(std) = 여러 제안서(고객)에 나온 묶음.
# 실행: ..\..\.venv\Scripts\python -X utf8 regroup.py   → store/out/meta.json · d1_seed.sql · facets.json 다시 씀
#       (store_build.py 끝에서도 finish(meta) 로 부름)
import collections
import json
import re
import sys
from multiprocessing import Pool
from pathlib import Path

from PIL import Image

APP = Path(__file__).resolve().parent.parent
OUT = APP / "store" / "out"
HEADS = OUT / "heads.jsonl"

STD_MIN_DOCS = 4  # 고객(제안서) 4곳 이상에 나온 묶음 = 자주 쓰는 표준 장표
SEP = re.compile(r"\t|\s{3,}|[\v\n]|\s*[|｜│›»]\s*|\s+>\s+")
LEAD = re.compile(r"^(\s*(\[[^\]]{1,14}\]|【[^】]{1,14}】|<[^>]{1,14}>|\d+(\.\d+)*\.?\)?|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.?|[IVX]{1,4}\.|[①-⑳]|[A-Z][\.\)]|\(\d+\)|[-–·•▶■□◆◇○●]))+\s*")
TAIL = re.compile(r"(\s*(\(계속\)|\(cont[^)]*\)|\(\d+\s*/\s*\d+\)|\*+|\(\s*\d+\s*\)|[-–:：]))+$", re.I)
# 같은 뜻 다른 이름(자주 나오는 것만) — 키는 띄어쓰기 · 기호 뺀 소문자
SYN = {
    "추진일정계획": "추진일정", "프로젝트일정": "추진일정", "프로젝트추진일정": "추진일정", "세부추진일정": "추진일정",
    "추진일정및계획": "추진일정", "수행일정": "추진일정", "프로젝트수행일정": "추진일정", "일정계획": "추진일정",
    "목적및범위": "프로젝트목적및범위", "프로젝트목적및과업범위": "프로젝트목적및범위", "프로젝트배경및목적": "프로젝트목적및범위",
    "수행조직및역할": "수행조직", "프로젝트추진조직": "수행조직", "프로젝트수행조직": "수행조직", "추진조직": "수행조직",
    "추진조직및역할": "수행조직", "프로젝트조직": "수행조직",
    "hcg특장점": "제안특장점", "제안특징및장점": "whyhcg",
    "목차": "목차", "contents": "목차", "보상simulation": "simulation", "보상시뮬레이션": "simulation",
}
# 흔한 말이라 이름만으로는 다른 내용끼리 섞이는 것 → 윗칸(대분류)과 붙여 이름을 지음(예: 테스트 방안 · 개요)
GENERIC = {"개요", "overview", "process", "프로세스", "절차", "framework", "방법론", "approach", "방안", "구성", "원칙", "기준", "체계",
           "운영", "사례", "예시", "output", "결과", "시사점", "summary", "요약", "방법", "내용", "주요내용", "세부내용", "구성요소"}


# ---------- 어깨끈 뽑기(장표 파일 → heads.jsonl, 한 번 뽑으면 다시 안 함) ----------
def _runs_text(tf):
    """제목 글. 굵기가 「상속 → 굵게 아님」으로 바뀌는 곳에 탭(대분류/소분류 경계)."""
    out = []
    for p in tf.paragraphs:
        prev = None
        for r in p.runs:
            b = r.font.bold
            if prev is not None and prev is not False and b is False:
                out.append("\t")
            out.append(r.text)
            prev = b
        out.append("\n")
    return "".join(out).strip()


def _head_one(pp):
    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER
    try:
        prs = Presentation(str(pp))
        W, H = prs.slide_width, prs.slide_height
        s = prs.slides[0]
        title, body, top = "", [], []
        for sh in s.shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            y, h, x = (sh.top or 0) / H, (sh.height or 0) / H, (sh.left or 0) / W
            if sh.is_placeholder:
                t = sh.placeholder_format.type
                if t in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) and not title:
                    title = _runs_text(sh.text_frame)
                    continue
                if t in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SUBTITLE) and y < 0.25 and not body:
                    body = [" ".join(p.text.split()) for p in sh.text_frame.paragraphs if p.text.strip()][:3]
                    continue
            if y < 0.10 and h < 0.12:
                top.append({"x": round(x, 3), "y": round(y, 3), "name": sh.name, "t": _runs_text(sh.text_frame)[:160]})
        return {"id": pp.stem, "title": title, "body": body, "top": top}
    except Exception as e:  # noqa: BLE001 — 깨진 파일은 제목 없이
        return {"id": pp.stem, "title": "", "body": [], "top": [], "err": repr(e)}


def load_heads(ids):
    have = {}
    if HEADS.exists():
        for line in open(HEADS, encoding="utf-8"):
            d = json.loads(line)
            have[d["id"]] = d
    need = [OUT / "pptx" / f"{i}.pptx" for i in ids if i not in have and (OUT / "pptx" / f"{i}.pptx").exists()]
    if need:
        print("어깨끈 뽑기", len(need), flush=True)
        with Pool(6) as pool:
            for d in pool.imap(_head_one, need, chunksize=20):
                have[d["id"]] = d
        with open(HEADS, "w", encoding="utf-8") as f:
            for d in have.values():
                f.write(json.dumps(d, ensure_ascii=False) + "\n")
    return have


# ---------- 이름 ----------
def clean(seg):
    seg = " ".join(seg.split())
    for _ in range(3):
        seg = re.sub(r"^[.,·:;\-–&]+\s*", "", seg)
        seg = LEAD.sub("", seg).strip()
        seg = TAIL.sub("", seg).strip()
    return seg


CLIENT_RX = None  # finish() 가 고객명 사전으로 채움


def client_rx(names):
    """고객명 떼기 — 「AJ네트웍스의 가설적 이슈」 → 「가설적 이슈」, 「코스맥스에 대한 이해」 → 「고객사에 대한 이해」."""
    alt = "|".join(re.escape(n) for n in sorted({n for n in names if len(n) >= 2}, key=len, reverse=True))
    return re.compile(r"(?:\(주\)|㈜)?\s*(?:%s)\s*(?:社|\(주\)|㈜)?(?:의\s*|(?P<p>에\s*대한|에게|와|과|을|를|은|는|이|가|에)(?=\s|$)|(?![가-힣A-Za-z0-9])\s*)" % alt)


def strip_client(seg):
    if CLIENT_RX is None:
        return seg
    seg = CLIENT_RX.sub(lambda m: ("고객사" + m.group("p") + " ") if m.group("p") else " ", seg)
    seg = re.sub(r"\(주\)|㈜", "", seg)
    return clean(seg)


def segments(raw, names):
    out = []
    for s in SEP.split(raw or ""):
        c = strip_client(clean(s))
        if len(c) < 2 or c in names or len(c) > 60:
            continue
        out.append(c)
    return out


def pick_head(h, names, old_title):
    """어깨끈 칸 목록. 제목 개체 틀 → 위쪽 「제목」 이름 상자 → 맨 위 왼쪽 짧은 글 → 지난 제목."""
    segs = segments(h.get("title", ""), names)
    if not segs:
        tops = sorted(h.get("top", []), key=lambda t: (not re.search(r"제목|title", t["name"], re.I), t["y"], t["x"]))
        for t in tops:
            if len(t["t"]) <= 80 and t["y"] <= 0.06:
                segs = segments(t["t"], names)
                if segs:
                    break
    if not segs:
        segs = segments(old_title, names)
    if not segs and h.get("body"):
        segs = [clean(h["body"][0])[:40]] if len(clean(h["body"][0])) >= 2 else []
    return segs


def key_of(name):
    k = re.sub(r"[^0-9a-z가-힣]", "", name.lower().replace("現", "현").replace("&", "및"))
    return SYN.get(k, k)


# ---------- 모양 비교 ----------
def dhash(path):
    try:
        im = Image.open(path).convert("L").resize((17, 16), Image.LANCZOS)
    except Exception:  # noqa: BLE001
        return None
    px = im.tobytes()
    v = 0
    for y in range(16):
        for x in range(16):
            v = (v << 1) | (px[y * 17 + x] > px[y * 17 + x + 1])
    return v


def toks(text):
    return set(t for t in re.findall(r"[0-9A-Za-z가-힣]{2,}", (text or "").lower()))


def jac(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def same_look(a, b, H, T):
    """같은 모양 — dHash 256비트 거리(무작위 두 장 중앙값 ~100, 1% ~74). 38 이하 = 같은 틀에 상자 수 · 글만 다름."""
    ha, hb = H.get(a), H.get(b)
    d = bin(ha ^ hb).count("1") if ha is not None and hb is not None else 256
    j = jac(T[a], T[b])
    big = len(T[a]) >= 8 and len(T[b]) >= 8
    return d <= 38 or (d <= 48 and j >= 0.4) or (big and j >= 0.8)


def doc_key(m):
    # 초안 · 최종본은 한 제안서로(같은 고객 · 같은 해). 고객 없는 HCG 자료는 파일 이름에서 날짜 · final 뺀 것
    if m.get("client"):
        return f"{m['client']}|{m.get('year', '')}"
    return re.sub(r"[_\s-]*(final|최종|v\d+|\d{6,8}).*$", "", m.get("doc", ""), flags=re.I)


TYPING = APP.parent / "work" / "typing"


def load_types():
    """AI 가 붙인 유형(work/typing/out_*.tsv: id · 유형 코드 · 대표 이름 · 확신) + 유형 사전(admin/types.json)."""
    tj = json.load(open(Path(__file__).with_name("types.json"), encoding="utf-8"))
    info = {t["code"]: t for t in tj["types"]}
    got = {}
    for f in sorted(TYPING.glob("out_*.tsv")):
        for line in open(f, encoding="utf-8"):
            p = line.rstrip("\n").split("\t")
            if len(p) == 4 and p[1] in info:
                got[p[0]] = (p[1], p[2].strip(), p[3].strip())
    return got, info


def rank(m):
    return (-(m["tier"] == 1), -bool(m["star"]), -(int(m["year"]) if str(m.get("year", "")).isdigit() else 0), m["k"])


def finish(meta, names=None, dry=False):
    if names is None:
        import store_build as B
        names = set(B.known_names())
    names = set(names)
    # 고객명 사전 + 장표마다 알아 둔 원본 고객 · 이전 고객명
    global CLIENT_RX
    allnames = names | {m["client"] for m in meta if m.get("client")} | {o for m in meta for o in (m.get("olds") or [])}
    allnames = {n for n in allnames if len(n) >= 2 and n.upper() != "HCG"}
    CLIENT_RX = client_rx(allnames)
    by = {m["id"]: m for m in meta}
    heads = load_heads(list(by))
    # 1) 이름 · 묶음 키
    for m in meta:
        segs = pick_head(heads.get(m["id"], {}), names, m.get("title", ""))
        if not segs:
            segs = ["(제목 없음)"]
        m["name"] = segs[-1][:80]
        m["head"] = " › ".join(segs)[:160]
        k = key_of(m["name"])
        if k in GENERIC and len(segs) >= 2:
            m["name"] = f"{segs[-2]} · {segs[-1]}"[:80]
            k = key_of(segs[-2]) + k
        m["fam"] = k if k and m["name"] != "(제목 없음)" else "_" + m["id"]
        m["kind"], m["kind_name"] = "", ""  # 예전 글자 검색 6종은 틀린 것이 많아 뺌(묶음으로 대신)
    TY, TINFO = load_types()
    typed = 0
    for m in meta:
        t = TY.get(m["id"])
        if not t:
            m["type"], m["chap"], m["tconf"] = "", "", ""
            continue
        code, nm, conf = t
        nm = strip_client(clean(nm)) or m["name"]
        m["type"], m["chap"], m["tconf"] = code, TINFO[code]["ch"], conf
        m["name"] = nm[:80]
        m["fam"] = code
        typed += 1
    # 2) 그림 · 글
    H = {m["id"]: dhash(OUT / "thumb" / f"{m['id']}.jpg") for m in meta}
    T = {m["id"]: toks(m["name"] + " " + m.get("text", "")) for m in meta}
    fams = collections.defaultdict(list)
    for m in meta:
        fams[m["fam"]].append(m["id"])
    # 3) 혼자인 장이 다른 묶음 장과 거의 같으면(제목만 조금 다름) 그 묶음으로
    import numpy as np
    big = [i for ids in fams.values() if len(ids) >= 2 for i in ids if H[i] is not None]
    hb = lambda v: np.frombuffer(v.to_bytes(32, "big"), dtype=np.uint8)  # noqa: E731 — 256비트 → 32바이트
    arr = np.stack([hb(H[i]) for i in big]) if big else np.zeros((0, 32), np.uint8)
    moved = 0
    for f, ids in list(fams.items()):
        if f in TINFO or len(ids) != 1 or H[ids[0]] is None or not len(arr):
            continue
        i = ids[0]
        d = np.unpackbits(arr ^ hb(H[i]), axis=1).sum(axis=1)
        best = None
        for x in np.argsort(d)[:8]:
            if d[x] > 8:
                break
            if jac(T[i], T[big[x]]) >= 0.6:
                best = (int(d[x]), big[x])
                break
        if best:
            nf = by[best[1]]["fam"]
            by[i]["fam"] = nf
            fams[nf].append(i)
            del fams[f]
            moved += 1
    # 4) 묶음 안에서 모양 나누기
    for f, ids in fams.items():
        # 대표 장(좋은 순)에 붙이기 — 사슬처럼 이어져 다른 모양까지 묶이는 것을 막음
        leaders, groups = [], {}
        for i in sorted(ids, key=lambda i: rank(by[i])):
            for ld in leaders:
                if same_look(ld, i, H, T):
                    groups[ld].append(i)
                    break
            else:
                leaders.append(i)
                groups[i] = [i]
        docs = {doc_key(by[i]) for i in ids}
        names_c = collections.Counter(by[i]["name"] for i in ids)
        fam_name = TINFO[f]["name"] if f in TINFO else names_c.most_common(1)[0][0]
        vars_ = sorted(groups.values(), key=lambda g: (-len({doc_key(by[i]) for i in g}), min(rank(by[i]) for i in g)))
        for vi, g in enumerate(vars_):
            g.sort(key=lambda i: rank(by[i]))
            for n, i in enumerate(g):
                m = by[i]
                m["var"] = f"{g[0]}"
                m["var_n"] = len(g)
                m["var_rep"] = 1 if n == 0 else 0
                m["var_ord"] = vi
        fid = f if f in TINFO else f"f{min(by[i]['k'] for i in ids):05d}"
        for i in ids:
            m = by[i]
            m["fam"] = fid
            m["fam_name"] = fam_name
            m["fam_n"] = len(ids)
            m["fam_v"] = len(vars_)
            m["fam_docs"] = len(docs)
    # 5) 자주 쓰는 표준 장표(목차 · 간지 · 표지 · 가격은 빼고)
    for ids in fams.values():  # 묶음 단위로 정함(한 묶음이 두 목록에 나뉘지 않게)
        ms = [by[i] for i in ids]
        cover = sum(m["p"] in ("P21", "P26") for m in ms) * 2 >= len(ms)
        if ms[0].get("type"):  # 유형 = 표지 · 목차 · 간지만 빼고 모두 Library 유형 카드
            s = 0 if ms[0]["type"] == "T601" else 1
        else:
            s = 1 if (ms[0]["fam_docs"] >= STD_MIN_DOCS and not cover and ms[0]["name"] != "(제목 없음)") else 0
        for m in ms:
            m["std"] = s
    import store_build as B
    for m in meta:
        m["action"] = B.action_of(m["p"], m["s"], "")
    nf = len({m["fam"] for m in meta})
    nstd = len({m["fam"] for m in meta if m["std"]})
    print(f"유형 붙은 장 {typed} · 묶음 {nf} · Layout {len({m['var'] for m in meta})} · 유형 카드 {nstd} · 혼자 → 묶음 {moved}")
    if not dry:
        write(meta)
    return meta


COLS = ["id", "k", "title", "text", "olds", "mods", "mod", "p", "p_name", "s", "s_name", "grade", "star", "kind", "kind_name", "action",
        "client", "industry", "year", "doc", "page", "bytes", "tier",
        "name", "head", "fam", "fam_name", "fam_n", "fam_v", "fam_docs", "var", "var_n", "var_rep", "var_ord", "std", "type", "chap", "tconf"]


def write(meta):
    json.dump(meta, open(OUT / "meta.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    with open(OUT / "d1_seed.sql", "w", encoding="utf-8") as f:
        f.write("DELETE FROM slides;\n")
        for m in meta:
            vals = []
            for c in COLS:
                v = m.get(c, "")
                if isinstance(v, list):
                    v = "|" + "|".join(v) + "|"
                if isinstance(v, bool):
                    v = int(v)
                vals.append(str(v) if isinstance(v, int) else "'" + str(v).replace("'", "''") + "'")
            f.write(f"INSERT INTO slides ({','.join(COLS)}) VALUES ({','.join(vals)});\n")
    fc = {"mods": collections.Counter(), "purposes": collections.Counter(), "structs": collections.Counter(), "industries": collections.Counter()}
    for m in meta:
        for x in m["mods"]:
            fc["mods"][x] += 1
        fc["purposes"][(m["p"], m["p_name"])] += 1
        fc["structs"][(m["s"], m["s_name"])] += 1
        fc["industries"][m["industry"]] += 1
    facets = {"total": len(meta), "fams": len({m["fam"] for m in meta}), "std_fams": len({m["fam"] for m in meta if m["std"]}),
              "mods": dict(fc["mods"]),
              "purposes": [{"p": k[0], "p_name": k[1], "n": n} for k, n in fc["purposes"].most_common()],
              "structs": [{"s": k[0], "s_name": k[1], "n": n} for k, n in fc["structs"].most_common()],
              "industries": [{"industry": k, "n": n} for k, n in fc["industries"].most_common()],
              "tiers": dict(collections.Counter(m["tier"] for m in meta)),
              "chaps": dict(collections.Counter(m.get("chap") for m in meta if m.get("chap")))}
    json.dump(facets, open(OUT / "facets.json", "w", encoding="utf-8"), ensure_ascii=False)


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    dry = "--dry" in sys.argv
    meta = finish(json.load(open(OUT / "meta.json", encoding="utf-8")), dry=dry)
    if dry:  # 확인용 — 결과를 따로 씀
        json.dump(meta, open(OUT / "meta_dry.json", "w", encoding="utf-8"), ensure_ascii=False)
