# Draft 만들기의 기계 부분: extract(고칠 글 목록 뽑기) · apply(Claude 답을 덱에 넣기 + 어깨끈 · 핵심 메시지 + 서식 정리)
# 글 안 「**…**」 = 원문에서 굵게(또는 다른 서식)였던 구절. Claude 가 같은 성격의 구절에 다시 붙이면 그 서식으로 넣는다.
import copy, json, os, re, sys, hashlib, collections, time
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn

def stype(sh):
    """도형 종류. python-pptx 가 모르는 도형(특수 자동 도형)은 None."""
    try:
        return sh.shape_type
    except NotImplementedError:
        return None

D = os.path.dirname(os.path.abspath(__file__))  # 작업 폴더(hcgd.py 가 바꿈)
HERE = D  # media_index.json 이 있는 폴더
CLIENT = "TCK"  # 고객 약칭(hcgd.py 가 바꿈)
OUT = ""  # meta.json 이 있는 폴더(hcgd.py 가 바꿈)
SKIP = re.compile(r"^(본 제안서 내의 컨텐츠는|\d{1,3}$|HCG$)")
KEEP_ROLES = ("소개",)  # 그대로 넣는 장(고쳐 쓰지 않고 서식도 건드리지 않음)


def walk(shapes):
    for sh in shapes:
        if stype(sh) == MSO_SHAPE_TYPE.GROUP:
            yield from walk(sh.shapes)
        else:
            yield sh


def ph_kind(sh):
    """제목(어깨끈) = "title", 핵심 메시지 자리 = "msg", 그 밖 = None."""
    if not sh.is_placeholder:
        return None
    pf = sh.placeholder_format
    t = str(pf.type)
    if "TITLE" in t and "SUB" not in t and "CENTER" not in t:
        return "title"
    if "BODY" in t and pf.idx == 10:
        return "msg"
    return None


def load_order():
    return json.load(open(os.path.join(D, "order.json"), encoding="utf-8"))


# ---------------- 문단 글 ↔ 서식 ----------------
def lname(el):
    return el.tag.split("}")[1] if isinstance(el.tag, str) else ""


def _lst_rpr(el):
    """조각이 든 문단 수준의 도형 lstStyle defRPr 들(lvlN → defPPr 순)."""
    p = el.getparent()
    tb = p.getparent() if p is not None else None
    lst = tb.find(qn("a:lstStyle")) if tb is not None else None
    if lst is None:
        return []
    ppr = p.find(qn("a:pPr"))
    lv = int(ppr.get("lvl", "0")) + 1 if ppr is not None else 1
    out = []
    for tag in (f"a:lvl{lv}pPr", "a:defPPr"):
        x = lst.find(qn(tag))
        if x is not None and x.find(qn("a:defRPr")) is not None:
            out.append(x.find(qn("a:defRPr")))
    return out


def skey(el):
    """글 조각의 눈에 보이는 서식 열쇠(굵게 · 기울임 · 밑줄 · 색). 값이 없으면 도형 목록 서식 → 기본값으로 풀어 비교(「0」과 「없음」을 같게)."""
    rpr = el.find(qn("a:rPr"))
    chain = ([rpr] if rpr is not None else []) + _lst_rpr(el)

    def val(attr, default):
        return next((x.get(attr) for x in chain if x.get(attr) is not None), default)
    fill = next((x.find(qn("a:solidFill")) for x in chain if x.find(qn("a:solidFill")) is not None), None)
    col = re.sub(r"\s+xmlns:\w+=\"[^\"]*\"", "", etree.tostring(fill, encoding="unicode")) if fill is not None else ""
    b = val("b", "0")
    return ("1" if b in ("1", "true") else "0", val("i", "0"), val("u", "none"), col)


def segs(para):
    out = []
    for el in para._p:
        t = lname(el)
        if t in ("r", "fld"):
            out.append((el.findtext(qn("a:t")) or "", skey(el), el))
        elif t == "br":
            out.append(("\n", None, el))
    return out


def plan_styles(para):
    """(바탕 조각, 강조 조각, 머리말 통일?) — 바탕 = 글자 수가 가장 많은 서식. 짧은 한 줄 머리말에 굵게가 섞였으면 굵게로 통일."""
    ss = [(t, k, el) for t, k, el in segs(para) if k is not None]
    if not ss:
        return None, None, False
    cnt = collections.Counter()
    for t, k, _ in ss:
        cnt[k] += len(t.strip())
    base_k = max(cnt, key=lambda k: (cnt[k], -[x[1] for x in ss].index(k)))
    plain = "".join(t for t, _, _ in segs(para)).strip()
    others = [k for k in cnt if k != base_k and cnt[k] > 0]
    unify = False
    if others and len(plain) <= 20 and "\n" not in plain:
        bold = [k for k in cnt if k[0] == "1" and cnt[k] > 0]
        if bold:
            base_k, others, unify = bold[0], [], True
    base_el = next(el for t, k, el in ss if k == base_k)
    emph_el = next((el for t, k, el in ss if k in others), None)
    return base_el, emph_el, unify


def ptext(para, marks=True):
    """문단 글. 문단 안 줄바꿈(a:br)은 \\n. marks=True 면 바탕과 서식이 다른 구절을 **…** 로 감쌈."""
    sg = segs(para)
    if not marks:
        txt = "".join(t for t, _, _ in sg)
        return re.sub(r"\n+$", "", txt.replace("\x0b", "\n")).strip(" ")
    base_el, emph_el, unify = plan_styles(para)
    base_k = skey(base_el) if base_el is not None else None
    flags = [None if (k is None or not t.strip() or unify or base_k is None) else (k != base_k) for t, k, _ in sg]
    for n, f in enumerate(flags):  # 줄바꿈 · 빈칸은 앞뒤가 모두 강조일 때만 강조 안으로
        if f is None:
            prev = next((x for x in reversed(flags[:n]) if x is not None), False)
            nxt = next((x for x in flags[n + 1:] if x is not None), False)
            flags[n] = "in" if (prev is True and nxt is True) else "out"
    out, on = [], False
    for (t, _, _), f in zip(sg, flags):
        want = f is True or f == "in"
        if want != on:
            out.append("**"); on = want
        out.append(t)
    if on:
        out.append("**")
    txt = "".join(out).replace("\x0b", "\n")
    txt = re.sub(r"\*\*(\s*)\*\*", r"\1", txt)  # 빈 강조
    return re.sub(r"\n+$", "", txt).strip(" ")


def plain(text):
    return text.replace("**", "")


def set_para(para, text):
    """문단 글 바꾸기: 바탕 서식으로 줄마다 조각을 만들고 줄 사이에 a:br. **…** 구절은 원문 강조 서식(없으면 바탕 + 굵게)."""
    p = para._p
    base_el, emph_el, unify = plan_styles(para)
    if base_el is None:
        return False
    brpr = base_el.find(qn("a:rPr"))
    erpr = emph_el.find(qn("a:rPr")) if emph_el is not None else None
    if erpr is None:
        erpr = copy.deepcopy(brpr) if brpr is not None else p.makeelement(qn("a:rPr"), {})
        erpr.set("b", "1")
    for el in [el for el in p if lname(el) in ("r", "br", "fld")]:
        p.remove(el)
    end = p.find(qn("a:endParaRPr"))
    pieces = text.split("**")
    if len(pieces) % 2 == 0:  # 짝이 안 맞으면 강조 없이
        pieces = [text.replace("**", "")]
    new = []
    for n, piece in enumerate(pieces):
        rpr = erpr if n % 2 else brpr
        for m, line in enumerate(piece.split("\n")):
            if m:
                br = p.makeelement(qn("a:br"), {})
                if brpr is not None:
                    br.append(copy.deepcopy(brpr))
                new.append(br)
            if not line:
                continue
            r = p.makeelement(qn("a:r"), {})
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            t = r.makeelement(qn("a:t"), {})
            t.text = line
            r.append(t)
            new.append(r)
    for el in new:
        if end is not None:
            end.addprevious(el)
        else:
            p.append(el)
    return True


def set_place(sh, text, title=False):
    """어깨끈(제목 자리) · 핵심 메시지 자리를 레이아웃 서식 그대로 새로 씀.
    어깨끈은 「대분류    중분류    소분류」(공백 4칸으로 나눔) — 대분류는 레이아웃대로(굵게), 중분류는 굵게 뺌, 소분류는 굵게 빼고 기울임."""
    if not title:  # 핵심 메시지 칸 = 2줄 높이(실제 제안서와 같게 46.4pt). 본문은 보통 115pt 아래에서 시작
        try:
            l, t, w = sh.left, sh.top, sh.width
            if l is not None and t is not None and w:
                sh.left, sh.top, sh.width, sh.height = l, t, w, Pt(46.4)
        except Exception:
            pass
    tb = sh.text_frame._txBody
    ps = tb.findall(qn("a:p"))
    p0 = ps[0]
    for p in ps[1:]:
        tb.remove(p)
    for el in [el for el in p0 if lname(el) in ("r", "br", "fld", "endParaRPr")]:
        p0.remove(el)
    text = re.sub(r"[ \t]+$", "", plain(text or ""))
    parts = [(text, None, None)]
    if title:
        bits = [b for b in re.split(r"\s{3,}", text.strip(), maxsplit=2) if b]
        parts = [(b + ("    " if k < len(bits) - 1 else ""), None if k == 0 else "0", "1" if k == 2 else None) for k, b in enumerate(bits)] or [("", None, None)]
    for t, b, it in parts:
        r = p0.makeelement(qn("a:r"), {})
        rpr = r.makeelement(qn("a:rPr"), {"lang": "ko-KR", "altLang": "en-US"})
        if b is not None:
            rpr.set("b", b)
        if it is not None:
            rpr.set("i", it)
        r.append(rpr)
        te = r.makeelement(qn("a:t"), {}); te.text = t; r.append(te)
        p0.append(r)


# ---------------- 뽑기 ----------------
def extract(deck, order_path, plan_path):
    prs = Presentation(deck)
    order = json.load(open(order_path, encoding="utf-8"))
    plan = {p["no"]: p for p in json.load(open(plan_path, encoding="utf-8"))}
    meta = {r["id"]: r for r in json.load(open(os.path.join(OUT, "meta.json"), encoding="utf-8"))}
    H = prs.slide_height
    slides = []
    for i, (s, o) in enumerate(zip(prs.slides, order), 1):
        info = {"slide": i, "role": o["role"], "no": o.get("no"), "items": [], "charts": []}
        if o["role"] in KEEP_ROLES:
            slides.append(info); continue
        if o.get("no"):
            p = plan[o["no"]]
            info["plan"] = {"head": p["head"], "msg": p["msg"], "type": p["type"]}
            src = meta.get(p["pick"], {})
            info["from"] = f"{src.get('client') or '기준 제안서'} {src.get('year') or ''}".strip()
        for sh in walk(s.shapes):
            if o.get("no") and ph_kind(sh):
                continue  # 어깨끈 · 핵심 메시지는 계획대로 따로 넣음
            if sh.name == "HCG_MODS":
                continue  # 모듈 간지의 모듈 목록(계획에서 만든 것)
            top = round((sh.top or 0) / H, 2) if sh.top is not None else 0
            if sh.has_text_frame:
                for k, para in enumerate(sh.text_frame.paragraphs):
                    t = ptext(para)
                    if t and not SKIP.match(plain(t)):
                        info["items"].append({"id": f"{i}:{sh.shape_id}:{k}", "y": top, "t": t})
            elif getattr(sh, "has_table", False) and sh.has_table:
                for r, row in enumerate(sh.table.rows):
                    for c, cell in enumerate(row.cells):
                        for k, para in enumerate(cell.text_frame.paragraphs):
                            t = ptext(para)
                            if t:
                                info["items"].append({"id": f"{i}:{sh.shape_id}:r{r}c{c}:{k}", "y": top, "t": t})
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                ch = sh.chart
                plots = list(ch.plots)
                cats = [str(c) for c in plots[0].categories] if plots else []
                sers = [{"name": se.name, "values": [None if v is None else round(v, 4) for v in se.values]} for pl in plots for se in pl.series]
                title = ch.chart_title.text_frame.text if ch.has_title and ch.chart_title.has_text_frame else ""
                info["charts"].append({"id": f"{i}:{sh.shape_id}:chart", "title": title, "categories": cats, "series": sers})
        slides.append(info)
    json.dump(slides, open(os.path.join(D, "items.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    n = sum(len(s["items"]) for s in slides)
    print("장", len(slides), "글 항목", n, "차트", sum(len(s["charts"]) for s in slides), "글자", sum(len(x["t"]) for s in slides for x in s["items"]))


# ---------------- 서식 맞추기 ----------------
HCG_FONTS = {"맑은 고딕", "Malgun Gothic", "맑은 고딕 Semilight", "Malgun Gothic Semilight"}
SYM = re.compile(r"Wingdings|Symbol|Webdings|Marlett", re.I)
KO = re.compile(r"[ᄀ-ᇿ㄰-㆏가-힣一-鿿㐀-䶿]")
EN = re.compile(r"[A-Za-z0-9]")
SPC = {"K": "-100", "E": "-30"}  # 한글 · 한자 −1pt, 영문 · 숫자 −0.3pt


def script_runs(t):
    """글을 한글 묶음(K) · 영문/숫자 묶음(E)으로 나눔. 띄어쓰기 · 기호는 앞 묶음에 붙임(맨 앞이면 뒤 묶음)."""
    out = []
    for ch in t:
        c = "K" if KO.match(ch) else ("E" if EN.match(ch) else None)
        if out and (c is None or c == out[-1][1]):
            out[-1][0] += ch
        elif out and out[-1][1] is None:
            out[-1][0] += ch; out[-1][1] = c
        else:
            out.append([ch, c])
    return [(s, c or "E") for s, c in out]


def tidy(prs, skip=()):
    """글꼴(HCG 외 → 맑은 고딕) · 자간(한글 −1pt · 영문 · 숫자 −0.3pt, 한 조각 안에 섞이면 나눔). 기호 글꼴 · 그대로 넣은 장은 건드리지 않음."""
    st = collections.Counter()
    for i, s in enumerate(prs.slides, 1):
        if i in skip:
            continue
        for r in list(s._element.iter(qn("a:r"))):
            t = r.findtext(qn("a:t")) or ""
            if not t.strip():
                continue
            rpr = r.find(qn("a:rPr"))
            if rpr is None:
                rpr = r.makeelement(qn("a:rPr"), {}); r.insert(0, rpr)
            sym = False
            for tag in ("a:latin", "a:ea", "a:sym"):
                el = rpr.find(qn(tag))
                tf = el.get("typeface") if el is not None else None
                if tf and SYM.search(tf):
                    sym = True
                elif tf and tag != "a:sym" and not tf.startswith("+") and tf not in HCG_FONTS:
                    el.set("typeface", "맑은 고딕"); st["글꼴"] += 1
            if sym:
                continue
            parts = script_runs(t)
            if len(parts) == 1:
                if rpr.get("spc") != SPC[parts[0][1]]:
                    rpr.set("spc", SPC[parts[0][1]]); st["자간"] += 1
                continue
            for txt, c in parts:  # 한 조각 안에 한글 · 영문이 섞임 → 나눠서 따로 자간
                nr = copy.deepcopy(r)
                nr.find(qn("a:t")).text = txt
                nr.find(qn("a:rPr")).set("spc", SPC[c])
                r.addprevious(nr)
            r.getparent().remove(r); st["자간 나눔"] += 1
    return st


# ---------------- 표지 업종 ----------------
IND = ""  # 고객 업종(hcgd.py 가 바꿈)
COVER_KIND = {"금융": "금융", "공공·연구": "금융", "지주·그룹": "금융", "기타": "금융",
              "자동차·기계·중공업": "제조", "전자·반도체": "제조", "화학·소재·에너지": "제조", "식품·소비재": "제조", "바이오·제약": "제조",
              "건설·부동산": "건설",
              "IT·플랫폼·게임": "서비스/IT", "유통·서비스·레저": "서비스/IT", "교육·미디어": "서비스/IT", "물류·운송": "서비스/IT"}
COVERS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "covers.json")


def cover_band(prs, ind):
    """표지 레이아웃 위쪽 그림 띠를 업종에 맞는 HCG 표준 표지(금융 · 제조 · 건설 · 서비스/IT, blue)로 바꿈. 바꾼 종류를 돌려줌."""
    kind = COVER_KIND.get(ind or "", "")
    if not kind or not os.path.exists(COVERS):
        return None
    band = json.load(open(COVERS, encoding="utf-8")).get(kind)
    lay = next((s.slide_layout for s in prs.slides if s.slide_layout.name.startswith("표지")), None)
    if not band or lay is None or lay.name == f"표지_{kind}_blue":
        return None
    from lxml import etree
    from pptx.opc.package import Part
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    import base64
    tree = lay.shapes._spTree
    old = [sh._element for sh in lay.shapes if stype(sh) == MSO_SHAPE_TYPE.GROUP and (sh.top or 0) < 2810000]
    if not old:
        return None
    at = list(tree).index(old[0])
    for el in old:
        tree.remove(el)
    pkg, rid_map = lay.part.package, {}
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    for b in band["shapes"]:
        el = etree.fromstring(b["xml"])
        for node in el.iter():
            for attr in (f"{{{R}}}embed", f"{{{R}}}link"):
                rid = node.get(attr)
                if rid and rid in band["parts"]:
                    if rid not in rid_map:
                        pt = band["parts"][rid]
                        part = Part(pkg.next_image_partname(pt["ext"]), pt["ct"], pkg, base64.b64decode(pt["b64"]))
                        rid_map[rid] = lay.part.relate_to(part, RT.IMAGE)
                    node.set(attr, rid_map[rid])
        tree.insert(at, el); at += 1
    lay._element.cSld.set("name", f"표지_{kind}_blue")
    return kind


# ---------------- 넣기 ----------------
def apply(deck, answers_glob, out):
    t0 = time.time()
    prs = Presentation(deck)
    texts, charts, checks = {}, {}, []
    import glob
    for f in sorted(glob.glob(os.path.join(D, answers_glob))):
        a = json.load(open(f, encoding="utf-8"))
        texts.update(a.get("texts", {})); charts.update(a.get("charts", {})); checks += a.get("checks", [])
    midx = json.load(open(os.path.join(HERE, "media_index.json"), encoding="utf-8"))
    order = load_order()
    plan = {p["no"]: p for p in json.load(open(os.path.join(D, "plan.json"), encoding="utf-8"))}
    st = collections.Counter()
    SH = prs.slide_height
    keep = {i for i, o in enumerate(order, 1) if o["role"] in KEEP_ROLES}
    try:
        if cover_band(prs, IND):
            st["표지 업종 그림"] += 1
    except Exception as e:  # 표지 그림은 못 바꿔도 나머지는 진행
        st["표지 업종 그림 실패"] += 1
        checks.append({"slide": 1, "tag": "표지 업종 그림", "note": f"표지 위 그림을 업종({IND})에 맞게 못 바꿈: {str(e)[:60]}"})
    # 표지 레이아웃 위쪽의 고객 로고(HCG 공용 그림이 아닌 것) → 표지 장에 「고객 약칭」 글 상자로 바꿈
    for i, (s, o) in enumerate(zip(prs.slides, order), 1):
        if o["role"] != "표지":
            continue
        for sh in list(s.slide_layout.shapes):
            if stype(sh) != MSO_SHAPE_TYPE.PICTURE or (sh.top or 0) > SH * 0.4:
                continue
            e = midx.get(hashlib.md5(sh.image.blob).hexdigest())
            if e is None or len([x for x in e["clients"] if x]) <= 1:
                tb = s.shapes.add_textbox(sh.left, sh.top, sh.width, sh.height)
                tb.text_frame.text = CLIENT
                r = tb.text_frame.paragraphs[0].runs[0]; r.font.bold = True; r.font.size = Emu(int(sh.height * 0.45))
                sh._element.getparent().remove(sh._element); st["표지 고객 로고 바꿈"] += 1
                checks.append({"slide": i, "id": f"{i}:{tb.shape_id}:0", "tag": "고객 로고 넣기", "note": f"표지 고객 로고 자리에 「{CLIENT}」 글 — 로고 파일로 바꾸기"})
    # 수행 조직 장의 고객 로고(여러 고객이 같이 쓰는 HCG 그림이 아닌 것) → 「고객 약칭」 글 상자
    for i, (s, o) in enumerate(zip(prs.slides, order), 1):
        pl = plan.get(o.get("no")) if o.get("no") else None
        if not pl or pl.get("type") != "T302":
            continue
        for sh in list(walk(s.shapes)):
            if stype(sh) != MSO_SHAPE_TYPE.PICTURE:
                continue
            e = midx.get(hashlib.md5(sh.image.blob).hexdigest())
            if e is not None and len([x for x in e["clients"] if x]) >= 5:
                continue  # HCG 로고 등 공용 그림
            if sh.width < sh.height * 1.3:
                continue  # 사람 · 아이콘 그림(로고는 가로로 긺)
            tb = s.shapes.add_textbox(sh.left, sh.top, sh.width, sh.height)
            tb.text_frame.text = CLIENT
            r = tb.text_frame.paragraphs[0].runs[0]; r.font.bold = True; r.font.size = Emu(int(min(sh.height * 0.55, Pt(20))))
            sh._element.getparent().remove(sh._element); st["수행 조직 고객 로고 바꿈"] += 1
            checks.append({"slide": i, "id": f"{i}:{tb.shape_id}:0", "tag": "고객 로고 넣기", "note": f"수행 조직 고객 로고 자리에 「{CLIENT}」 글 — 로고 파일로 바꾸기"})
    for i, (s, o) in enumerate(zip(prs.slides, order), 1):
        if i in keep:
            continue
        pl = plan.get(o.get("no")) if o.get("no") else None
        emptied = []
        for sh in list(walk(s.shapes)):
            kind = ph_kind(sh) if pl else None
            if kind and sh.has_text_frame:
                set_place(sh, pl["head"] if kind == "title" else pl["msg"], title=(kind == "title")); st["어깨끈 · 메시지"] += 1
                continue
            if sh.has_text_frame:
                had = bool(sh.text_frame.text.strip())
                for k, para in enumerate(list(sh.text_frame.paragraphs)):
                    key = f"{i}:{sh.shape_id}:{k}"
                    if key in texts:
                        v = texts[key]
                        if v == "" and len(sh.text_frame.paragraphs) > 1:
                            para._p.getparent().remove(para._p); st["문단 지움"] += 1
                        elif set_para(para, v):
                            st["글"] += 1
                if had and not sh.text_frame.text.strip():
                    emptied.append(sh)
            elif getattr(sh, "has_table", False) and sh.has_table:
                for r, row in enumerate(sh.table.rows):
                    for c, cell in enumerate(row.cells):
                        for k, para in enumerate(cell.text_frame.paragraphs):
                            key = f"{i}:{sh.shape_id}:r{r}c{c}:{k}"
                            if key in texts and set_para(para, texts[key]):
                                st["표 칸"] += 1
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                key = f"{i}:{sh.shape_id}:chart"
                if key in charts:
                    try:
                        blank = put_chart(sh, charts[key])
                    except Exception:  # 차트 하나 못 바꿔도 나머지는 진행
                        blank = True
                    if blank:
                        st["차트 빈 값"] += 1
                        checks.append({"slide": i, "id": f"{i}:{sh.shape_id}:chart", "tag": "차트 값 넣기", "note": f"차트 값 비움(공시에 없음) — 이 차트가 {CLIENT}에 맞는지 보고 값 넣기 또는 장 바꾸기", "kind": "차트"})
                    else:
                        st["차트"] += 1
        for sh in emptied:  # 글이 모두 비게 된 도형(목차 남는 칸 등)은 지움
            el = sh._element
            if el.getparent() is not None:
                el.getparent().remove(el); st["빈 도형 지움"] += 1
    # 다른 고객 그림은 지우지 않음 — 마무리에서 노란 딱지를 붙임
    st.update(tidy(prs, keep))
    prs.save(out)
    json.dump(checks, open(os.path.join(D, "checks.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    st["초"] = round(time.time() - t0, 1)
    print(json.dumps(st, ensure_ascii=False))


def put_chart(sh, c):
    """차트 값 바꾸기(차트 안 엑셀까지). 숫자 서식 · 고정 축 범위는 새 값 크기에 맞춰 유지. 빈 값이면 True."""
    cs = sh.chart._chartSpace
    if cs.xpath(".//c:scatterChart | .//c:bubbleChart"):
        return True  # 분산 · 거품 차트는 값 구조가 달라 그대로 두고 확인 딱지(값 넣기)
    old_fmt = [(x.xpath("./c:val//c:formatCode/text()") or ["General"])[0] for x in cs.iter(qn("c:ser"))]
    ax_old = collections.defaultdict(float)
    for pl in cs.find(qn("c:chart")).find(qn("c:plotArea")):
        ids = pl.xpath("./c:axId/@val")
        for ser in pl.findall(qn("c:ser")):
            for v in ser.xpath("./c:val//c:v/text()"):
                for a in ids:
                    try:
                        ax_old[a] = max(ax_old[a], abs(float(v)))
                    except ValueError:
                        pass
    cd = CategoryChartData()
    cd.categories = c["categories"]
    blank = True
    for n_se, se in enumerate(c["series"]):
        vals = se.get("values") or [None] * len(c["categories"])
        blank = blank and all(v is None for v in vals)
        cd.add_series(se["name"], vals, number_format=old_fmt[n_se] if n_se < len(old_fmt) else None)
    sh.chart.replace_data(cd)
    cs = sh.chart._chartSpace
    for pl in cs.find(qn("c:chart")).find(qn("c:plotArea")):
        ids = pl.xpath("./c:axId/@val")
        newmax = 0.0
        for ser in pl.findall(qn("c:ser")):
            for v in ser.xpath("./c:val//c:v/text()"):
                try:
                    newmax = max(newmax, abs(float(v)))
                except ValueError:
                    pass
        for ax in cs.iter(qn("c:valAx")):
            aid = ax.find(qn("c:axId")).get("val")
            if aid in ids and ax_old.get(aid) and newmax:
                f = newmax / ax_old[aid]
                for tag in ("c:max", "c:min"):
                    el = ax.find(qn("c:scaling")).find(qn(tag))
                    if el is not None:
                        el.set("val", repr(round(float(el.get("val")) * f, 6)))
    if c.get("title") and sh.chart.has_title:
        sh.chart.chart_title.text_frame.text = c["title"]
    return blank


if __name__ == "__main__":
    if sys.argv[1] == "extract":
        extract(*sys.argv[2:5])
    else:
        apply(*sys.argv[2:5])
