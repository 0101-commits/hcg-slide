# 합치기(PowerPoint 없이): 기준 틀 파일(표지 · 간지 · 끝장)을 바탕으로, 창고 한 장 파일들을 계획 순서대로 넣는다.
# 가져온 장은 모두 바탕 마스터의 「본문」 레이아웃에 붙인다(장마다 다른 마스터를 끌고 오지 않음).
# 그 대신 원본 파일에서 물려받던 서식이 바뀌지 않게, 원본과 바탕이 다른 것만 장에 직접 적어 둔다:
#   ① 기본 글 서식(presentation.xml defaultTextStyle — 굵게 · 글머리표 글꼴 · 줄 간격 · 글자 색 등)
#   ② 테마 색(accent 등) — 원본 색을 RGB 로 박음
# 사용: merge.py order.json out.pptx frame.pptx   (order = [{"base": 7} | {"lib": "s01234"}, ...])
import copy, json, os, re, sys, zipfile, hashlib, posixpath, time
from lxml import etree

LIB = os.environ.get("HCG_LIB_PPTX", "")  # 창고 장표 파일 폴더(hcgd.py 가 내려받은 곳으로 바꿈)
STD_LAYOUT = "본문"  # 가져온 장을 붙일 바탕 레이아웃 이름
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
RT_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
RT_NOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
RT_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
RT_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"


def rels_name(part):
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def resolve(part, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def relpath(frm, to):
    return posixpath.relpath(to, posixpath.dirname(frm))


def xml(el):
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


class Pkg:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.parts = {n: z.read(n) for n in z.namelist()}
        ct = etree.fromstring(self.parts["[Content_Types].xml"])
        self.defaults = {e.get("Extension").lower(): e.get("ContentType") for e in ct.findall("ct:Default", NS)}
        self.overrides = {e.get("PartName").lstrip("/"): e.get("ContentType") for e in ct.findall("ct:Override", NS)}

    def ctype(self, part):
        return self.overrides.get(part) or self.defaults.get(part.rsplit(".", 1)[-1].lower())

    def rels(self, part):
        rn = rels_name(part)
        if rn not in self.parts:
            return []
        x = etree.fromstring(self.parts[rn])
        return [(e.get("Id"), e.get("Type"), e.get("Target"), e.get("TargetMode")) for e in x.findall("pr:Relationship", NS)]

    def rel(self, part, rtype):
        return next((resolve(part, tg) for _, t, tg, m in self.rels(part) if t == rtype and m != "External"), None)

    def reach(self, start, skip_types=()):
        seen, todo = [], [start]
        while todo:
            p = todo.pop()
            if p in seen or p not in self.parts:
                continue
            seen.append(p)
            for _, t, tg, mode in self.rels(p):
                if mode == "External" or t in skip_types:
                    continue
                todo.append(resolve(p, tg))
        return seen


# ---------------- 테마 색 ----------------
SLOTS = ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]


def scheme_colors(pkg, master):
    """마스터 기준으로 schemeClr 이름 → 실제 RGB(clrMap 거쳐 테마 색)."""
    mx = etree.fromstring(pkg.parts[master])
    cmap = mx.find(P + "clrMap")
    th = pkg.rel(master, RT_THEME)
    if th is None or cmap is None:
        return {}
    tx = etree.fromstring(pkg.parts[th])
    cs = tx.find(f".//{A}clrScheme")
    slot = {}
    for el in cs:
        c = el[0] if len(el) else None
        if c is not None:
            slot[el.tag.split("}")[1]] = (c.get("val") if c.tag == A + "srgbClr" else c.get("lastClr") or "").upper()
    out = {}
    for name in SLOTS:
        out[name] = slot.get(cmap.get(name, name), "")
    for k in ("dk1", "lt1", "dk2", "lt2"):
        out[k] = slot.get(k, "")
    return out


def bake_colors(root, diff):
    """원본과 바탕의 테마 색이 다른 이름만 schemeClr → srgbClr(원본 색). lumMod 같은 변형은 그대로 둠."""
    n = 0
    for el in list(root.iter(A + "schemeClr")):
        v = el.get("val")
        if v in diff and diff[v]:
            el.tag = A + "srgbClr"
            el.set("val", diff[v])
            n += 1
    return n


# ---------------- 기본 글 서식 ----------------
PPR_ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont",
             "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst"]
PPR_GROUP = {"buClrTx": "buClr", "buClr": "buClr", "buSzTx": "buSz", "buSzPct": "buSz", "buSzPts": "buSz", "buFontTx": "buFont",
             "buFont": "buFont", "buNone": "buType", "buAutoNum": "buType", "buChar": "buType", "buBlip": "buType"}
RPR_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag", "highlight",
             "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
RPR_GROUP = {"noFill": "fill", "solidFill": "fill", "gradFill": "fill", "blipFill": "fill", "pattFill": "fill", "grpFill": "fill",
             "effectLst": "effect", "effectDag": "effect", "uLnTx": "uLn", "uLn": "uLn", "uFillTx": "uFill", "uFill": "uFill"}
SKIP_PPR = {"defRPr", "extLst", "tabLst"}
SKIP_RA = {"lang", "altLang", "dirty", "err", "noProof", "smtClean", "smtId", "bmk"}
# 아무 데서도 정하지 않았을 때 PowerPoint 기본값
DEF_PA = {"marL": "0", "indent": "0", "algn": "l", "rtl": "0", "eaLnBrk": "1", "fontAlgn": "auto", "latinLnBrk": "0", "hangingPunct": "1"}
DEF_RA = {"kumimoji": "0", "b": "0", "i": "0", "u": "none", "strike": "noStrike", "kern": "0", "cap": "none", "spc": "0", "baseline": "0", "sz": "1800"}
DEF_PC = {"lnSpc": '<a:lnSpc xmlns:a="%s"><a:spcPct val="100000"/></a:lnSpc>',
          "spcBef": '<a:spcBef xmlns:a="%s"><a:spcPts val="0"/></a:spcBef>',
          "spcAft": '<a:spcAft xmlns:a="%s"><a:spcPts val="0"/></a:spcAft>',
          "buClr": '<a:buClrTx xmlns:a="%s"/>', "buSz": '<a:buSzTx xmlns:a="%s"/>', "buFont": '<a:buFontTx xmlns:a="%s"/>',
          "buType": '<a:buNone xmlns:a="%s"/>'}
DEF_RC = {"fill": '<a:solidFill xmlns:a="%s"><a:schemeClr val="tx1"/></a:solidFill>',
          "latin": '<a:latin xmlns:a="%s" typeface="+mn-lt"/>', "ea": '<a:ea xmlns:a="%s" typeface="+mn-ea"/>', "cs": '<a:cs xmlns:a="%s" typeface="+mn-cs"/>'}


def lname(el):
    return el.tag.split("}")[1] if isinstance(el.tag, str) else ""


def canon(el):
    if el is None:
        return None
    s = etree.tostring(el, encoding="unicode")
    return re.sub(r'\s+xmlns:\w+="[^"]*"', "", s)


class Lvl:
    """한 수준(lvlN)의 문단 · 글자 속성 묶음."""
    def __init__(self):
        self.pa, self.pc, self.ra, self.rc = {}, {}, {}, {}

    def add(self, ppr):
        if ppr is None:
            return
        self.pa.update({k: v for k, v in ppr.attrib.items() if k != "lvl"})
        for ch in ppr:
            n = lname(ch)
            if n in SKIP_PPR or not n:
                continue
            self.pc[PPR_GROUP.get(n, n)] = ch
        d = ppr.find(A + "defRPr")
        if d is not None:
            self.ra.update({k: v for k, v in d.attrib.items() if k not in SKIP_RA})
            for ch in d:
                n = lname(ch)
                if n:
                    self.rc[RPR_GROUP.get(n, n)] = ch


def styles(lst):
    """defaultTextStyle · lstStyle 요소 → 수준 1~9 별 Lvl(defPPr 위에 lvlN 을 덮음)."""
    out = {}
    for lv in range(1, 10):
        L = Lvl()
        if lst is not None:
            L.add(lst.find(A + "defPPr"))
            L.add(lst.find(A + f"lvl{lv}pPr"))
        out[lv] = L
    return out


def put_child(parent, new, order):
    n = lname(new)
    idx = order.index(n) if n in order else len(order)
    for i, ch in enumerate(parent):
        cn = lname(ch)
        if cn in order and order.index(cn) > idx:
            ch.addprevious(new)
            return
    parent.append(new)


def has_group(parent, group, gmap):
    return any(gmap.get(lname(ch), lname(ch)) == group for ch in parent)


def bake_text(root, src, base):
    """자리표시자 아닌 글 상자 · 표 칸의 문단마다, 원본과 바탕 기본값이 다른 속성을 직접 적음(그 자리에서 정한 값은 그대로)."""
    diffs = {}
    for lv in range(1, 10):
        s, b = src[lv], base[lv]
        pa = {k: s.pa.get(k, DEF_PA.get(k)) for k in set(s.pa) | set(b.pa) if s.pa.get(k, DEF_PA.get(k)) != b.pa.get(k, DEF_PA.get(k))}
        ra = {k: s.ra.get(k, DEF_RA.get(k)) for k in set(s.ra) | set(b.ra) if s.ra.get(k, DEF_RA.get(k)) != b.ra.get(k, DEF_RA.get(k))}
        pc, rc = {}, {}
        for g in set(s.pc) | set(b.pc):
            se = canon(s.pc.get(g)) or (DEF_PC.get(g, "") % NS["a"] if g in DEF_PC else None)
            be = canon(b.pc.get(g)) or (DEF_PC.get(g, "") % NS["a"] if g in DEF_PC else None)
            if se and se != be:
                pc[g] = s.pc.get(g) if g in s.pc else etree.fromstring(DEF_PC[g] % NS["a"])
        for g in set(s.rc) | set(b.rc):
            se = canon(s.rc.get(g)) or (DEF_RC.get(g, "") % NS["a"] if g in DEF_RC else None)
            be = canon(b.rc.get(g)) or (DEF_RC.get(g, "") % NS["a"] if g in DEF_RC else None)
            if se and se != be:
                rc[g] = s.rc.get(g) if g in s.rc else etree.fromstring(DEF_RC[g] % NS["a"])
        diffs[lv] = (pa, pc, ra, rc)
    if not any(any(d) for d in diffs.values()):
        return 0
    bodies = []
    for sp in root.iter(P + "sp"):
        if sp.find(f"{P}nvSpPr/{P}nvPr/{P}ph") is not None:
            continue
        tb = sp.find(P + "txBody")
        if tb is not None:
            bodies.append(tb)
    bodies += list(root.iter(A + "tc"))  # 표 칸(a:tc/a:txBody)
    n = 0
    for tb in bodies:
        if lname(tb) == "tc":
            tb = tb.find(A + "txBody")
            if tb is None:
                continue
        local = styles(tb.find(A + "lstStyle"))
        for p in tb.findall(A + "p"):
            ppr = p.find(A + "pPr")
            lv = int(ppr.get("lvl", "0")) + 1 if ppr is not None else 1
            pa, pc, ra, rc = diffs.get(lv, diffs[1])
            loc = local[lv]
            if pa or pc:
                if ppr is None:
                    ppr = etree.SubElement(p, A + "pPr"); p.remove(ppr); p.insert(0, ppr)
                for k, v in pa.items():
                    if k not in loc.pa and ppr.get(k) is None and v is not None:
                        ppr.set(k, v); n += 1
                for g, el in pc.items():
                    if g not in loc.pc and not has_group(ppr, g, PPR_GROUP):
                        put_child(ppr, copy.deepcopy(el), PPR_ORDER); n += 1
            if not (ra or rc):
                continue
            targets = []
            for el in p:
                t = lname(el)
                if t in ("r", "fld"):
                    rpr = el.find(A + "rPr")
                    if rpr is None:
                        rpr = etree.Element(A + "rPr"); el.insert(0, rpr)
                    targets.append(rpr)
                elif t == "br" and el.find(A + "rPr") is not None:
                    targets.append(el.find(A + "rPr"))
                elif t == "endParaRPr":
                    targets.append(el)
            for rpr in targets:
                for k, v in ra.items():
                    if k not in loc.ra and rpr.get(k) is None and v is not None:
                        rpr.set(k, v); n += 1
                for g, el in rc.items():
                    if g not in loc.rc and not has_group(rpr, g, RPR_GROUP):
                        put_child(rpr, copy.deepcopy(el), RPR_ORDER); n += 1
    return n


def fit_placeholders(root):
    """제목(어깨끈) · 핵심 메시지 자리는 바탕 레이아웃 자리 · 크기를 따르게(장에 박힌 위치 지움). 그 밖 자리표시자는 보통 도형으로."""
    n = 0
    for sp in root.iter(P + "sp"):
        ph = sp.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
        if ph is None:
            continue
        if ph.get("type") == "title" or (ph.get("type", "body") == "body" and ph.get("idx") == "10"):
            spr = sp.find(P + "spPr")
            x = spr.find(A + "xfrm") if spr is not None else None
            if x is not None:
                spr.remove(x); n += 1
        else:
            ph.getparent().remove(ph)
    return n


def main(order_path, out_path, base_path):
    t0 = time.time()
    order = json.load(open(order_path, encoding="utf-8"))
    B = Pkg(base_path)
    pres = etree.fromstring(B.parts["ppt/presentation.xml"])
    prels = etree.fromstring(B.parts["ppt/_rels/presentation.xml.rels"])
    rid_target = {e.get("Id"): resolve("ppt/presentation.xml", e.get("Target")) for e in prels.findall("pr:Relationship", NS)}
    sldIdLst = pres.find("p:sldIdLst", NS)
    base_slides = [rid_target[e.get("{%s}id" % NS["r"])] for e in sldIdLst.findall("p:sldId", NS)]
    mst_lst = pres.find("p:sldMasterIdLst", NS)
    base_master = rid_target[mst_lst.find("p:sldMasterId", NS).get("{%s}id" % NS["r"])]
    std_layout = None
    for _, t, tg, _m in B.rels(base_master):
        if t == RT_LAYOUT:
            lp = resolve(base_master, tg)
            if etree.fromstring(B.parts[lp]).find("p:cSld", NS).get("name") == STD_LAYOUT:
                std_layout = lp
    if std_layout is None:
        raise SystemExit(f"바탕 틀에 「{STD_LAYOUT}」 레이아웃이 없음")
    base_style = styles(pres.find("p:defaultTextStyle", NS))
    base_colors = scheme_colors(B, base_master)
    rid_n = [max(int(re.sub(r"\D", "", k) or 0) for k in rid_target) + 1]
    media_by_hash = {hashlib.md5(v).hexdigest(): k for k, v in B.parts.items() if k.startswith("ppt/media/")}

    def fresh(name):
        d, f = posixpath.split(name)
        m = re.match(r"([^\d]*?)(\d*)(\.[^.]+)$", f)
        stem, ext = (m.group(1), m.group(3)) if m else (f, "")
        k = 1
        while True:
            cand = posixpath.join(d, f"{stem}{k}{ext}")
            if cand not in B.parts:
                return cand
            k += 1

    stats = {"lib": 0, "media_reused": 0, "baked_text": 0, "baked_color": 0, "placeholders": 0}
    final = []
    for item in order:
        if "base" in item:
            sp = base_slides[item["base"] - 1]
            if sp in final:  # 같은 틀 장을 또 쓰면(모듈 간지) 장 부품을 복제 — 메모 장 연결은 뺌
                np_ = fresh(sp)
                B.parts[np_] = B.parts[sp]
                B.overrides[np_] = B.overrides.get(sp)
                if rels_name(sp) in B.parts:
                    rx = etree.fromstring(B.parts[rels_name(sp)])
                    for e in list(rx.findall("pr:Relationship", NS)):
                        if e.get("Type") == RT_NOTES:
                            rx.remove(e)
                    B.parts[rels_name(np_)] = xml(rx)
                sp = np_
            final.append(sp); continue
        L = Pkg(os.path.join(LIB, item["lib"] + ".pptx"))
        lp = etree.fromstring(L.parts["ppt/presentation.xml"])
        lrel = {e.get("Id"): resolve("ppt/presentation.xml", e.get("Target")) for e in etree.fromstring(L.parts["ppt/_rels/presentation.xml.rels"]).findall("pr:Relationship", NS)}
        src_slide = lrel[lp.find("p:sldIdLst/p:sldId", NS).get("{%s}id" % NS["r"])]
        src_layout = L.rel(src_slide, RT_LAYOUT)
        src_master = L.rel(src_layout, RT_MASTER) if src_layout else None
        src_style = styles(lp.find("p:defaultTextStyle", NS))
        sc = scheme_colors(L, src_master) if src_master else {}
        cdiff = {k: v for k, v in sc.items() if v and base_colors.get(k) and v != base_colors[k]}
        parts = L.reach(src_slide, skip_types=(RT_NOTES, RT_LAYOUT))  # 장 자신의 부품(그림 · 차트 · 엑셀 등)만
        mapping = {}
        for p in parts:
            if p.endswith(".rels"):
                continue
            if p.startswith("ppt/media/"):
                h = hashlib.md5(L.parts[p]).hexdigest()
                if h in media_by_hash:
                    mapping[p] = media_by_hash[h]; stats["media_reused"] += 1; continue
            np_ = fresh(p)
            B.parts[np_] = b""  # 자리 잡기
            mapping[p] = np_
            if p.startswith("ppt/media/"):
                media_by_hash[hashlib.md5(L.parts[p]).hexdigest()] = np_
        for p, np_ in mapping.items():
            if p not in parts or B.parts.get(np_) != b"":
                continue
            data = L.parts[p]
            if p.endswith(".xml") and (p == src_slide or p.startswith(("ppt/charts/", "ppt/diagrams/"))):
                root = etree.fromstring(data)
                if p == src_slide:
                    stats["baked_text"] += bake_text(root, src_style, base_style)
                    stats["placeholders"] += fit_placeholders(root)
                if cdiff:  # 글 서식을 박은 뒤(박은 글자 색도 함께 바꾸게)
                    stats["baked_color"] += bake_colors(root, cdiff)
                data = xml(root)
            B.parts[np_] = data
            ct = L.ctype(p)
            ext = np_.rsplit(".", 1)[-1].lower()
            if ext not in B.defaults or B.defaults[ext] != ct:
                B.overrides[np_] = ct
            rn = rels_name(p)
            if rn in L.parts:
                rx = etree.fromstring(L.parts[rn])
                for e in list(rx.findall("pr:Relationship", NS)):
                    if e.get("TargetMode") == "External":
                        continue
                    if e.get("Type") == RT_NOTES:
                        rx.remove(e); continue
                    if e.get("Type") == RT_LAYOUT:
                        e.set("Target", relpath(np_, std_layout)); continue
                    abs_t = resolve(p, e.get("Target"))
                    if abs_t in mapping:
                        e.set("Target", relpath(np_, mapping[abs_t]))
                B.parts[rels_name(np_)] = xml(rx)
        final.append(mapping[src_slide]); stats["lib"] += 1
    # 장 목록 다시 쓰기
    for e in list(sldIdLst):
        sldIdLst.remove(e)
    for e in list(prels.findall("pr:Relationship", NS)):
        if e.get("Type") == RT_SLIDE:
            prels.remove(e)
    for k, sp in enumerate(final):
        rid = f"rId{rid_n[0]}"; rid_n[0] += 1
        etree.SubElement(prels, "{%s}Relationship" % NS["pr"], Id=rid, Type=RT_SLIDE, Target=relpath("ppt/presentation.xml", sp))
        se = etree.SubElement(sldIdLst, "{%s}sldId" % NS["p"]); se.set("id", str(256 + k)); se.set("{%s}id" % NS["r"], rid)
    # 구역(section) · 사용자 지정 쇼는 옛 장 id 를 가리키므로 지움. 보기 설정의 개요 목록도 옛 장을 붙잡으므로 지움
    for el in list(pres.iter("{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst")):
        ext = el.getparent(); ext.getparent().remove(ext)
    cs = pres.find("p:custShowLst", NS)
    if cs is not None:
        pres.remove(cs)
    if "ppt/viewProps.xml" in B.parts:
        vx = etree.fromstring(B.parts["ppt/viewProps.xml"])
        for el in list(vx.findall("p:outlineViewPr", NS)):
            vx.remove(el)
        B.parts["ppt/viewProps.xml"] = xml(vx)
        B.parts.pop("ppt/_rels/viewProps.xml.rels", None)
    B.parts["ppt/presentation.xml"] = xml(pres)
    B.parts["ppt/_rels/presentation.xml.rels"] = xml(prels)
    # 닿지 않는 부품 지우기
    keep = {"[Content_Types].xml", "_rels/.rels"}
    todo = [tg.lstrip("/") for _, t, tg, mode in [(e.get("Id"), e.get("Type"), e.get("Target"), e.get("TargetMode")) for e in etree.fromstring(B.parts["_rels/.rels"]).findall("pr:Relationship", NS)] if mode != "External"]
    while todo:
        p = todo.pop()
        if p in keep or p not in B.parts:
            continue
        keep.add(p)
        rn = rels_name(p)
        if rn in B.parts:
            keep.add(rn)
            for _, t, tg, mode in B.rels(p):
                if mode != "External":
                    todo.append(resolve(p, tg))
    dropped = [n for n in B.parts if n not in keep]
    for n in dropped:
        del B.parts[n]
        B.overrides.pop(n, None)
    ct = etree.Element("{%s}Types" % NS["ct"], nsmap={None: NS["ct"]})
    for ext, c in B.defaults.items():
        etree.SubElement(ct, "{%s}Default" % NS["ct"], Extension=ext, ContentType=c)
    for n, c in B.overrides.items():
        if n in B.parts and c:
            etree.SubElement(ct, "{%s}Override" % NS["ct"], PartName="/" + n, ContentType=c)
    B.parts["[Content_Types].xml"] = xml(ct)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", B.parts.pop("[Content_Types].xml"))
        for n, v in B.parts.items():
            z.writestr(n, v)
    stats.update({"slides": len(final), "dropped_parts": len(dropped), "sec": round(time.time() - t0, 1), "MB": round(os.path.getsize(out_path) / 1e6, 1)})
    print(json.dumps(stats, ensure_ascii=False))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
