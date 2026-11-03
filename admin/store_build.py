# 장표 창고 묶음 만들기 — 표준 장표 후보(재사용 「상」 ∪ 별표)에서 골라 한 장짜리 PPTX · 미리보기 · 검색 목록을 만든다.
# 실행: ..\..\.venv\Scripts\python -X utf8 store_build.py --limit 50        (맛보기 50장)
#       ..\..\.venv\Scripts\python -X utf8 store_build.py --all             (973장 전부, 2단계)
# 결과: app/store/out/{pptx,img,thumb}/<id>.* + meta.json + d1_seed.sql
# 원본(OneDrive 02 Proposal · 08 Reference)은 읽기만 한다.
import argparse, collections, copy, json, os, re, sys, time
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

import scrub

HERE = Path(__file__).resolve().parent
APP = HERE.parent
WORK = APP.parent / "work"
OUT = APP / "store" / "out"
ROOT_PROP = Path(r"C:\Users\cgpar\OneDrive - 휴먼컨설팅그룹\02 Proposal")
ROOT_REF = Path(r"C:\Users\cgpar\OneDrive - 휴먼컨설팅그룹\08 Reference\01 제안서")

PURPOSE = dict(re.findall(r"(P\d\d) (.+?)(?= · P\d\d|$)", "P01 여는 종합 · P02 환경 변화·트렌드 · P03 고객 이해 · P04 가설적 이슈 · P05 목적·범위·전체 접근 · P06 Why HCG·차별점 · P07 방법론·프로세스 · P08 진단 방법·틀 · P09 진단 결과·시사점 · P10 벤치마킹·사례 참고 · P11 대안·옵션 비교 · P12 As-Is→To-Be · P13 제도 설계 상세 · P14 시뮬레이션·효과 · P15 운영 절차·역할 · P16 로드맵·실행·변화관리 · P17 일정 · P18 산출물 이미지 · P19 솔루션·시스템 · P20 정의·개념 · P21 목차·간지·표지·끝장 · P22 조직·투입 인력 · P23 회사 소개·실적 · P24 법·제도·판례 · P25 관리 방안 · P26 가격·견적 · P99 기타"))
STRUCT = dict(re.findall(r"([A-Z]\d?) (.+?)(?= · [A-Z]\d? |$)", "L1 좌우 2단 · L2 3~4단 기둥 · L6 줄 구조 · L3 카드 격자 · L4 위 요약+아래 상세 · L5 좌우 비교 · F1 가로 단계 · F2 세로 단계 · F3 순환 · F4 로드맵·타임라인 · F5 간트 · F6 깔때기 · F7 계단 · H1 계층·조직도 · H2 로직 트리 · H3 피라미드 · H4 하우스 · H5 프레임워크 블록 · H6 중심-방사 · H7 벤·동심원 · C1 2×2 매트릭스 · C2 옵션 비교표 · C3 평가표 · C4 장단점 · D1 차트+시사점 · D2 숫자 카드 · D3 분포·밴드 · D4 산식 · T1 일반 표 · T2 정의서·양식 · T3 역할표 · T4 체크리스트 · P1 로고 사례 · P2 인물·이력 카드 · P3 인용·말풍선 · P4 화면 캡처 · P5 사진·이미지 · G 글만"))

# HR 모듈 — 장표 제목·글의 낱말로 대략 나눔(2단계에서 에이전트가 다시 붙임)
MODS = [
    ("평가", r"평가|성과\s*관리|KPI|MBO|OKR|다면|Calibration|캘리브레이션|목표\s*(수립|관리)|Check-?in|BSC"),
    ("보상", r"보상|연봉|성과급|임금|Pay|인센티브|수당|호봉|급여|처우|Profit Sharing|Target Incentive|PS\s*vs|재원|Payout|Band"),
    ("육성", r"리더십|육성|교육|경력\s*개발|CDP|코칭|승계|핵심\s*인재|HiPo|역량\s*개발"),
    ("직무", r"직무|Job\s|직군|직렬"),
    ("직급", r"직급|승진|호칭|직위|직책"),
    ("조직", r"조직\s*(설계|진단|구조|개편|운영|체계)|R&R|거버넌스|Span|조직도"),
    ("문화", r"조직\s*문화|문화|핵심\s*가치|몰입|Engagement|소통|서베이"),
    ("인력", r"적정\s*인력|인력\s*(계획|운영|구조|분포|산정)|정원|인원\s*(구조|분포)|요원"),
    ("시스템", r"talenx|휴넬|HRIS|시스템|솔루션|플랫폼|DNapp|JaDE"),
    ("채용", r"채용|선발|면접|온보딩"),
    ("노사", r"노조|노사|노무|통상\s*임금|판례|근로\s*기준|단체\s*협약|임금\s*피크"),
]
COMMON_P = {"P06", "P17", "P21", "P22", "P23", "P25", "P26"}
MOD_LABEL = {"평가": "평가·성과관리", "보상": "보상", "육성": "리더십·육성", "직무": "직무", "직급": "직급·승진", "조직": "조직 설계·진단",
             "문화": "조직문화", "인력": "인력·적정인력", "시스템": "HR 시스템·솔루션", "채용": "채용", "노사": "노사·노무", "공통": "공통"}

NAMES = []
# 자주 찾는 표준 장표 6종(기획서 4절) — 넣은 뒤 할 일
STD = {"talenx": ("talenx 소개", "그대로"), "PSvsTI": ("PS vs TI 성과급", "고객명만"), "PEST": ("PEST 환경 분석", "칸 글"),
       "수행계획": ("수행계획 · 일정", "막대 · 주차"), "인력구조": ("인력구조", "숫자 · 차트"), "보상시뮬": ("보상 시뮬레이션", "숫자 · 차트")}


def mods_of(text, p):
    m = [k for k, rx in MODS if re.search(rx, text, re.I)]
    if p in COMMON_P or not m:
        m.append("공통")
    return m


def action_of(p, s, kind):
    if kind:
        return STD[kind][1]
    if p in ("P21", "P23", "P19", "P06"):
        return "그대로"
    if p == "P17" or s in ("F5", "F4"):
        return "막대 · 주차"
    if p == "P14" or (s or "").startswith("D"):
        return "숫자 · 차트"
    return "칸 글"


def client_year(name):
    base = re.split(r"[\\/]", name)
    proj, fn = (base[0], base[-1]) if len(base) > 1 else ("", base[-1])
    year = ""
    m = re.match(r"(20\d\d)\.\d\d", proj)
    if m:
        year = m.group(1)
    else:
        m = re.search(r"(?<!\d)(20[12]\d)(?:[01]\d[0-3]\d)?(?!\d)", fn) or re.search(r"(?<!\d)([12]\d)[01]\d[0-3]\d(?!\d)", fn)
        if m:
            y = m.group(1)
            year = y if len(y) == 4 else "20" + y
    client = ""
    if proj:
        parts = proj.split("_")
        client = parts[1] if len(parts) > 1 else ""
    else:
        m = re.search(r"\][\s_]*([^_\]\[]+?)_", fn)
        if m:
            client = m.group(1).strip()
    client = re.sub(r"\(.*?\)|\s+(장기계약|제안서)$", "", client).strip()
    if re.search(r"HCG|Company|휴먼컨설팅|Ghost|Time Saver|템플릿|Graphic|^@", client + fn[:3], re.I) and not proj:
        client = ""
    return client, year


def known_names():
    """이전 고객명 사전 — 스킬의 client_names.txt(관리자 스크립트로 만듦) + 제안 폴더 이름의 고객명."""
    names = set()
    f = APP / "skills" / "hcg-check" / "references" / "client_names.txt"
    if f.exists():
        names |= {l.strip() for l in open(f, encoding="utf-8") if l.strip() and not l.startswith("#")}
    for p in json.load(open(WORK / "projects.json", encoding="utf-8")) if (WORK / "projects.json").exists() else []:
        nm = p.get("project") or p.get("name") or "" if isinstance(p, dict) else str(p)
        c, _ = client_year(nm + "/x")
        if c:
            names.add(c)
    dm = WORK / "retag" / "docs_map.json"
    if dm.exists():
        names |= {v[0] for k, v in json.load(open(dm, encoding="utf-8")).items() if not k.startswith("_") and v[0]}
    return sorted({n for n in names if len(n) >= 2}, key=len, reverse=True)


def load_retag():
    """2단계 다시 붙이기 결과 — work/retag/out_*.tsv(sid → HR 모듈 코드) · docs_map.json(문서 → 고객 · 업종 · 연도)."""
    rt, dm = {}, {}
    for f in sorted((WORK / "retag").glob("out_*.tsv")):
        for l in open(f, encoding="utf-8"):
            c = l.rstrip("\n").split("\t")
            if len(c) >= 2 and c[0].startswith("s") and c[0] != "sid":
                codes = [x.strip() for x in c[1].split("/") if x.strip() in MOD_LABEL]
                if codes:
                    rt[c[0]] = codes
    f = WORK / "retag" / "docs_map.json"
    if f.exists():
        dm = {k: v for k, v in json.load(open(f, encoding="utf-8")).items() if not k.startswith("_")}
    return rt, dm


def load_index():
    idx = {}
    for l in open(WORK / "index.jsonl", encoding="utf-8"):
        r = json.loads(l)
        idx[f"{r['pid']}_{r['fid']}"] = r
    conv = {}
    if (WORK / "converted.jsonl").exists():
        for l in open(WORK / "converted.jsonl", encoding="utf-8"):
            r = json.loads(l)
            conv[f"{r['pid']}_{r['fid']}"] = r.get("conv")
    refs = {f["id"]: f for f in json.load(open(WORK / "ref" / "files.json", encoding="utf-8"))}
    return idx, conv, refs


def source_path(row, idx, conv, refs):
    if row["src"] == "ref":
        f = refs.get(row["doc"])
        if not f:
            return None
        if f["ext"] == ".ppt":  # convert_ppt.ps1 로 바꾼 사본
            c = WORK / "ref" / "conv" / f"{row['doc']}.pptx"
            return c if c.exists() else None
        return ROOT_REF / f["rel"]
    r = idx.get(row["doc"])
    if not r:
        return None
    if r["ext"] == ".pptx":
        return ROOT_PROP / r["project"] / r["rel"]
    c = conv.get(row["doc"])
    return Path(c) if c and Path(c).exists() else None


def _rel_part(part, rid):
    try:
        return part.related_part(rid)
    except KeyError:  # 원본 파일에 끊긴 관계가 있음
        return None


def _drop(part, rid):
    if rid in part.rels:
        part.drop_rel(rid)


def cut_slide(src, i, out):
    """원본에서 i번째(1부터) 장만 남긴 PPTX를 out에 저장. 나머지 장 · 다른 장으로 가는 링크 · 구역 목록 정리."""
    prs = Presentation(str(src))
    lst = prs.slides._sldIdLst
    ids = list(lst)
    keep = ids[i - 1]
    keep_slide = prs.slides[i - 1]
    drop_ids = set()
    for s in ids:
        if s is keep:
            continue
        drop_ids.add(s.get("id"))
        prs.part.drop_rel(s.rId)
        lst.remove(s)
    # 남긴 장에서 다른 장으로 가는 관계(하이퍼링크 등) 끊기
    sp = keep_slide.part
    for rId, rel in list(sp.rels.items()):
        if rel.reltype == RT.SLIDE:
            hits = [el for el in sp._element.iter()
                    if any(v == rId and k.endswith("}id") for k, v in el.attrib.items())]
            for el in hits:
                if el.getparent() is not None:
                    el.getparent().remove(el)
            sp.drop_rel(rId)
    # 구역(section) · 사용자 지정 쇼에서 지운 장 번호 빼기
    pel = prs.part._element
    for el in list(pel.iter()):
        tag = el.tag.split("}")[-1]
        if tag == "sldId" and el.getparent() is not lst and el.get("id") in drop_ids:
            el.getparent().remove(el)
        elif tag == "custShowLst":
            el.getparent().remove(el)
    for el in list(pel.iter()):
        if el.tag.split("}")[-1] == "section":
            lst2 = [c for c in el.iter() if c.tag.split("}")[-1] == "sldIdLst"]
            if lst2 and len(lst2[0]) == 0:
                el.getparent().remove(el)
    # 남긴 장이 쓰지 않는 마스터 · 레이아웃 빼기(파일 크기)
    used_layout = keep_slide.slide_layout
    used_master = used_layout.slide_master
    mlst = pel.find("{http://schemas.openxmlformats.org/presentationml/2006/main}sldMasterIdLst")
    for mid in list(mlst):
        rid = mid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if _rel_part(prs.part, rid) is not used_master.part:
            mlst.remove(mid)
            _drop(prs.part, rid)
    mp = used_master.part
    llst = mp._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}sldLayoutIdLst")
    for lid in list(llst):
        rid = lid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if _rel_part(mp, rid) is not used_layout.part:
            llst.remove(lid)
            _drop(mp, rid)
    # 파일에 심은 글꼴 빼기 — 장표를 넣을 때 따라가지 않고 크기만 큼(수십 MB)
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for efl in pel.findall("{http://schemas.openxmlformats.org/presentationml/2006/main}embeddedFontLst"):
        rids = [e.get(R) for e in efl.iter() if e.get(R)]
        pel.remove(efl)
        for rid in rids:
            _drop(prs.part, rid)
    # 문서 속성의 작성자 · 회사 지우기(이전 고객 정보 남지 않게)
    cp = prs.core_properties
    cp.author = "HCG"
    cp.last_modified_by = "HCG"
    cp.title = ""
    cp.subject = ""
    cp.keywords = ""
    cp.comments = ""
    prs.save(str(out))
    return keep_slide


def slide_text(slide):
    out = []

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:  # 묶음
                walk(sh.shapes)
                continue
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    out.append(t)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        t = c.text.strip()
                        if t:
                            out.append(t)
    walk(slide.shapes)
    note = ""
    if slide.has_notes_slide:
        note = slide.notes_slide.notes_text_frame.text.strip() if slide.notes_slide.notes_text_frame else ""
    return " ".join(" ".join(out).split()), note


def pick(rows, limit, std_map):
    """맛보기: 표준 6종에서 먼저, 나머지는 별표 → 「상」 순으로 HR 모듈이 고르게."""
    chosen, per_doc = [], collections.Counter()
    for kind in STD:
        for r in std_map.get(kind, [])[:3]:
            if len([c for c in chosen if c.get("_kind") == kind]) >= 2:
                break
            r = dict(r, _kind=kind)
            chosen.append(r)
            per_doc[r["doc"]] += 1
    if limit <= 0:
        seen = {c["k"] for c in chosen}
        return chosen + [r for r in rows if r["k"] not in seen]
    by_mod = collections.defaultdict(list)
    for r in sorted(rows, key=lambda r: (not r["star"], r["re"] != "상", r["k"])):
        by_mod[r["_mods"][0]].append(r)
    seen = {c["k"] for c in chosen}
    while len(chosen) < limit:
        added = False
        for m in list(MOD_LABEL):
            q = by_mod.get(m, [])
            while q and (q[0]["k"] in seen or per_doc[q[0]["doc"]] >= 2):
                q.pop(0)
            if q and len(chosen) < limit:
                r = q.pop(0)
                chosen.append(r)
                seen.add(r["k"])
                per_doc[r["doc"]] += 1
                added = True
        if not added:
            break
    return chosen


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=50)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--force", action="store_true", help="이미 만든 장도 다시")
    a = ap.parse_args()
    limit = 0 if a.all else a.limit

    global NAMES
    NAMES = known_names()
    print("고객명 사전", len(NAMES))
    RETAG, DOCS = load_retag()
    print("다시 붙이기", len(RETAG), "장 · 문서", len(DOCS))
    idx, conv, refs = load_index()
    tags = [json.loads(l) for l in open(WORK / "ref" / "review" / "tags_all.jsonl", encoding="utf-8")]
    slides = {json.loads(l)["k"]: json.loads(l) for l in open(WORK / "ref" / "review" / "slides_all.jsonl", encoding="utf-8")}
    cand = []
    for r in tags:
        if r["rep"] != r["k"] or not (r["re"] == "상" or r["star"]):
            continue
        s = slides.get(r["k"], {})
        r["_text0"] = (r.get("title") or "") + " " + s.get("text", "")
        r["_mods"] = RETAG.get(f"s{r['k']:05d}") or mods_of(r["_text0"], r["p"])
        r["_src"] = source_path(r, idx, conv, refs)
        cand.append(r)
    print("후보", len(cand), "· 원본 파일 있음", sum(1 for r in cand if r["_src"]))

    # 표준 6종 — plan3/std_pick.json 의 그림 경로로 장표를 찾음
    std_map = {}
    sp = APP.parent / "plan3" / "std_pick.json"
    by_png = {r["png"]: r for r in tags}
    if sp.exists():
        for kind, v in json.load(open(sp, encoding="utf-8")).items():
            lst = []
            for p in v["pick"]:
                r = by_png.get(p["png"])
                if not r:
                    continue
                r = dict(r)
                s = slides.get(r["k"], {})
                r["_text0"] = (r.get("title") or "") + " " + s.get("text", "")
                r["_mods"] = RETAG.get(f"s{r['k']:05d}") or mods_of(r["_text0"], r["p"])
                r["_src"] = source_path(r, idx, conv, refs)
                if r["_src"]:
                    lst.append(r)
            std_map[kind] = lst
    sel = pick([r for r in cand if r["_src"]], limit, std_map)
    print("고름", len(sel), collections.Counter(r["_mods"][0] for r in sel))

    for d in ("pptx", "img", "thumb"):
        (OUT / d).mkdir(parents=True, exist_ok=True)
    meta, fails = [], []
    old = {}
    if (OUT / "meta.json").exists():
        old = {m["id"]: m for m in json.load(open(OUT / "meta.json", encoding="utf-8"))}
    t0 = time.time()
    for n, r in enumerate(sel, 1):
        sid = f"s{r['k']:05d}"
        kind = r.get("_kind", "")
        pp = OUT / "pptx" / f"{sid}.pptx"
        try:
            if a.force or not pp.exists() or sid not in old:
                slide = cut_slide(r["_src"], r["i"], pp)
                text, note = slide_text(slide)
                scrub.scrub_file(pp, NAMES)
            else:
                text, note = old[sid]["text"], old[sid].get("note", "")
            # 확인: 다시 열어 장 1개인지
            chk = Presentation(str(pp))
            assert len(chk.slides) == 1, f"장 수 {len(chk.slides)}"
            if a.force or not (OUT / "img" / f"{sid}.jpg").exists() or not (OUT / "thumb" / f"{sid}.jpg").exists():
                im = Image.open(r["png"]).convert("RGB")
                w = 960
                im.resize((w, round(im.height * w / im.width)), Image.LANCZOS).save(OUT / "img" / f"{sid}.jpg", quality=84, optimize=True)
                w = 360
                im.resize((w, round(im.height * w / im.width)), Image.LANCZOS).save(OUT / "thumb" / f"{sid}.jpg", quality=80, optimize=True)
            title = (r.get("title") or "").strip() or (STD[kind][0] if kind else "")
            client, year = client_year(r["name"])
            dc, industry, dy = DOCS.get(r["doc"], ["", "", ""])
            client, year = (dc or client), (dy or year)
            if r["doc"] in DOCS and not dc:
                client = ""  # HCG 자체 자료
            hay = title + " " + text + " " + note
            olds = [nm for nm in NAMES if nm in hay and not any(nm != o and nm in o for o in NAMES if o in hay)]
            if client and client in hay and client not in olds:
                olds.insert(0, client)
            meta.append({
                "id": sid, "k": r["k"], "title": " ".join(title.split())[:120], "text": text[:4000], "note": note[:1000],
                "mods": [MOD_LABEL[m] for m in r["_mods"]], "mod": MOD_LABEL[r["_mods"][0]],
                "p": r["p"], "p_name": PURPOSE.get(r["p"], ""), "s": r["s"], "s_name": STRUCT.get(r["s"], r["s"]),
                "grade": r["re"], "star": bool(r["star"]), "kind": kind, "kind_name": STD[kind][0] if kind else "",
                "action": action_of(r["p"], r["s"], kind), "client": client, "olds": olds[:8], "year": year, "industry": industry,
                "doc": re.split(r"[\\/]", r["name"])[-1], "src": r["src"], "page": r["i"],
                "bytes": pp.stat().st_size,
            })
            print(f"[{n}/{len(sel)}] {sid} {pp.stat().st_size // 1024}KB {title[:30]} · {client} {year} · {time.time() - t0:.0f}s", flush=True)
        except Exception as e:
            fails.append((sid, str(r["_src"]), repr(e)))
            print(f"[{n}/{len(sel)}] {sid} 실패 {e!r}", flush=True)
            if pp.exists():
                pp.unlink()
    json.dump(meta, open(OUT / "meta.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    with open(OUT / "d1_seed.sql", "w", encoding="utf-8") as f:
        f.write("DELETE FROM slides;\n")
        cols = ["id", "k", "title", "text", "olds", "mods", "mod", "p", "p_name", "s", "s_name", "grade", "star", "kind", "kind_name", "action", "client", "industry", "year", "doc", "page", "bytes"]
        for m in meta:
            vals = []
            for c in cols:
                v = m[c]
                if isinstance(v, list):
                    v = "|" + "|".join(v) + "|"
                if isinstance(v, bool):
                    v = int(v)
                vals.append(str(v) if isinstance(v, int) else "'" + str(v).replace("'", "''") + "'")
            f.write(f"INSERT INTO slides ({','.join(cols)}) VALUES ({','.join(vals)});\n")
    print("끝", len(meta), "장 · 실패", len(fails), f"· {time.time() - t0:.0f}s")
    for x in fails:
        print("  실패", *x)


if __name__ == "__main__":
    main()
