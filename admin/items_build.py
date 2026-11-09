# Item — Library 장표에서 자주 쓰는 도형 하나 · 작은 도형 묶음을 유형별로 모아, 지금 장에 그대로 끼워 넣을 수 있게 함
# 글은 ///(번호 배지는 숫자 그대로), 테마 색은 원래 색(sRGB)으로 바꿔 어느 파일에 넣어도 같은 색. 그림 · 표 · 차트 든 것은 뺌.
# 실행: ..\..\.venv\Scripts\python -X utf8 items_build.py [--no-render]
#   → pane/public/pane/items.json(+ items/*.png 미리보기) · store/items_sheet.jpg(확인용)
# 주의: 미리보기는 PowerPoint COM — 열린 PowerPoint 가 있으면 멈춤
import collections
import copy
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont

APP = Path(__file__).resolve().parent.parent
OUT = APP / "store" / "out"
PANE = APP / "pane" / "public" / "pane"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NSMAP = {"a": A[1:-1], "p": P[1:-1], "r": R[1:-1]}
PER = 12  # 유형마다 Item 수

KINDS = [  # (코드, 이름, 설명)
    ("step", "단계 화살표", "오각형 · 쉐브론 단계 상자(번호 붙은 것 포함)"),
    ("arrow", "큰 화살표", "흐름을 잇는 큰 화살표 · 삼각 화살표"),
    ("badge", "번호 · 배지", "①②③ 번호 원 · 작은 배지"),
    ("header", "머리띠 상자", "칸 · 표 위 머리 상자"),
    ("callout", "말풍선 · 메모", "시사점 말풍선 · 접힌 메모"),
    ("brace", "괄호 · 묶음", "여러 칸을 묶는 괄호"),
    ("bar", "결론 띠", "장 아래 결론 · 강조 띠"),
    ("box", "제목 상자", "번호 + 제목 상자 · 둥근 제목 상자"),
]
STEP = {"homePlate", "chevron"}
ARROW = {"rightArrow", "downArrow", "leftRightArrow", "upArrow", "leftArrow", "notchedRightArrow", "stripedRightArrow"}
CALL = {"wedgeRoundRectCallout", "wedgeRectCallout", "wedgeEllipseCallout", "foldedCorner", "cloudCallout"}
BRACE = {"leftBrace", "rightBrace", "leftBracket", "rightBracket"}


def theme_colors(z):
    """테마 색 이름 → RGB(장 하나짜리 파일의 테마 · 색 매핑)."""
    names = [n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
    m = {}
    if not names:
        return m
    root = etree.fromstring(z.read(sorted(names)[0]))
    cs = root.find(f".//{A}clrScheme")
    if cs is None:
        return m
    for el in cs:
        tag = etree.QName(el).localname
        c = el[0] if len(el) else None
        if c is None:
            continue
        v = c.get("val") if etree.QName(c).localname == "srgbClr" else c.get("lastClr")
        if v:
            m[tag] = v
    alias = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
    for k, v in alias.items():
        if v in m:
            m[k] = m[v]
    return m


def fix_colors(el, tc):
    """schemeClr → srgbClr(밝기 조절 lumMod 등은 그대로 둠 — PowerPoint 가 sRGB 에도 적용)."""
    for sc in list(el.iter(A + "schemeClr")):
        v = tc.get(sc.get("val"))
        if not v:
            continue
        new = etree.Element(A + "srgbClr", val=v)
        for ch in list(sc):
            new.append(ch)
        sc.getparent().replace(sc, new)


def sanitize(el):
    """다른 파일에 끼워도 깨지지 않게 — 관계(r:id) 쓰는 것 · 개체 틀 표시 · 확장 정보 · 도형 스타일 참조를 뺌."""
    for c in list(el.iter()):
        if c.getparent() is None:
            continue
        ln = etree.QName(c).localname
        if ln in ("extLst", "custDataLst", "hlinkClick", "hlinkHover", "ph") or any(k.startswith(R) for k in c.attrib):
            c.getparent().remove(c)
    for c in el.iter():
        for k in [k for k in c.attrib if k.startswith(R)]:
            del c.attrib[k]


def blank_text(el, keep_digits):
    for p in el.iter(A + "p"):
        runs = [c for c in p if c.tag in (A + "r", A + "fld", A + "br")]
        txt = "".join(t.text or "" for t in p.iter(A + "t")).strip()
        if not runs or not txt:
            continue
        first = next((c for c in runs if c.tag == A + "r"), None)
        for c in runs:
            if c is not first:
                p.remove(c)
        if first is None:
            continue
        t = first.find(A + "t")
        if keep_digits and re.fullmatch(r"[0-9①-⑳A-Za-z]{1,3}\.?", txt):
            t.text = txt
        else:
            t.text = "///"
    for c in el.iter():
        if c.tag.endswith("}cNvPr"):
            c.attrib.pop("descr", None)
            c.attrib.pop("title", None)
            c.set("name", "HCG Item")


def ext_of(el):
    x = el.find(f"{P}spPr/{A}xfrm") if el.tag == P + "sp" else el.find(f"{P}grpSpPr/{A}xfrm")
    if x is None:
        return None
    off, ext = x.find(A + "off"), x.find(A + "ext")
    if off is None or ext is None:
        return None
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


def text_of(el):
    return "".join(t.text or "" for t in el.iter(A + "t")).strip()


def visible(el):
    """도형 자체가 보이는지 — 흰색이 아닌 채우기 · 그라데이션 · 선(글 서식 안의 색은 안 셈)."""
    for sp in el.iter(P + "spPr"):
        for f in sp:
            ln = etree.QName(f).localname
            if ln == "gradFill":
                return True
            if ln == "solidFill":
                c = f.find(A + "srgbClr")
                if c is None or c.get("val", "").upper() not in ("FFFFFF",):
                    return True
            if ln == "ln" and f.find(A + "solidFill") is not None and f.find(A + "noFill") is None:
                return True
    for st in el.iter(P + "style"):  # 도형 스타일로 채우기 · 선이 정해진 경우
        if st.find(A + "fillRef") is not None and st.find(A + "fillRef").get("idx") not in ("0", None):
            return True
    return False


def kind_of(el, W, H):
    """도형 하나 · 작은 묶음의 Item 유형(없으면 None)."""
    e = ext_of(el)
    if not e or not visible(el):
        return None
    if el.find(f"{P}spPr/{A}xfrm") is not None and el.find(f"{P}spPr/{A}xfrm").get("rot") not in (None, "0") and el.tag == P + "sp":
        g0 = el.find(f"{P}spPr/{A}prstGeom")
        if g0 is not None and g0.get("prst") in BRACE:
            return None  # 돌린 괄호는 미리보기 · 자리 잡기가 어긋남
    x, y, w, h = e
    if w <= 0 or h <= 0 or w > W * 0.98:
        return None
    if el.tag == P + "grpSp":
        kids = [k for k in el if k.tag in (P + "sp", P + "grpSp", P + "pic", P + "graphicFrame", P + "cxnSp")]
        if not 2 <= len(kids) <= 4 or any(k.tag in (P + "pic", P + "graphicFrame", P + "grpSp") for k in kids):
            return None
        prs = [k.find(f"{P}spPr/{A}prstGeom") for k in kids if k.tag == P + "sp"]
        geoms = {g.get("prst") for g in prs if g is not None}
        if geoms & STEP and w < W * 0.45:
            return "step"
        if ("rect" in geoms or "roundRect" in geoms) and w < W * 0.5 and h < H * 0.2 and any(re.fullmatch(r"[0-9]{1,2}\.?", text_of(k)) for k in kids):
            return "box"
        return None
    if el.tag != P + "sp":
        return None
    if el.find(f".//{A}blipFill") is not None:
        return None
    g = el.find(f"{P}spPr/{A}prstGeom")
    cust = el.find(f"{P}spPr/{A}custGeom")
    prst = g.get("prst") if g is not None else ("custom" if cust is not None else "")
    t = text_of(el)
    if prst in STEP and w < W * 0.45:
        return "step"
    if prst in ARROW and (w > W * 0.04 or h > H * 0.06):
        return "arrow"
    if prst == "custom" and el.find(f"{P}spPr/{A}gradFill") is not None and not t and h > H * 0.15 and w < W * 0.15:
        return "arrow"  # 회색 그라데이션 큰 삼각 화살표
    if prst in ("ellipse", "flowChartConnector") and 0 < len(t) <= 3 and w < W * 0.08 and abs(w - h) < 0.25 * w:
        return "badge"
    if prst in ("round2SameRect", "snip2SameRect") and t and h < H * 0.12:
        return "header"
    if prst in CALL:
        return "callout"
    if prst in BRACE and (h > H * 0.15 or w > W * 0.15):
        return "brace"
    if prst in ("rect", "roundRect", "round2SameRect") and t and w > W * 0.6 and H * 0.05 < h < H * 0.14 and y > H * 0.55:
        return "bar"
    return None


def signature(el, kind):
    """비슷한 Item 은 하나만 — 모양 · 채우기 · 선 · 크기 비율로 지문."""
    xml = etree.tostring(el).decode("utf-8")
    geoms = tuple(sorted(set(re.findall(r'prst="(\w+)"', xml))))
    fills = tuple(sorted(set(re.findall(r'srgbClr val="(\w{6})"', xml)))[:4])
    grad = "grad" if "gradFill" in xml else ""
    x, y, w, h = ext_of(el)
    ratio = round(w / max(h, 1), 1)
    ratio = min(ratio, 12)
    return (kind, geoms, fills, grad, round(ratio * 2) / 2)


def main():
    no_render = "--no-render" in sys.argv
    meta = json.load(open(OUT / "meta.json", encoding="utf-8"))
    tier = {m["id"]: m for m in meta}
    seen = collections.defaultdict(lambda: {"n": 0, "docs": set(), "xml": None, "src": None, "wh": None})
    files = sorted((OUT / "pptx").glob("s*.pptx"))
    for k, f in enumerate(files, 1):
        if f.stem not in tier:
            continue
        try:
            z = zipfile.ZipFile(f)
            pres = etree.fromstring(z.read("ppt/presentation.xml"))
            sz = pres.find(f"{P}sldSz")
            W, H = int(sz.get("cx")), int(sz.get("cy"))
            sn = next(n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n))
            root = etree.fromstring(z.read(sn))
            tree = root.find(f"{P}cSld/{P}spTree")
            tc = None
            for el in tree:
                kind = kind_of(el, W, H)
                if not kind:
                    continue
                if tc is None:
                    tc = theme_colors(z)
                el2 = copy.deepcopy(el)
                sanitize(el2)
                fix_colors(el2, tc)
                blank_text(el2, keep_digits=kind in ("badge", "box", "step"))
                sig = signature(el2, kind)
                d = seen[sig]
                d["n"] += 1
                d["docs"].add(tier[f.stem]["doc"])
                if d["xml"] is None or (tier[f.stem]["tier"] == 1 and d.get("t") != 1):
                    d["xml"] = etree.tostring(el2).decode("utf-8")
                    d["src"] = f.stem
                    d["wh"] = ext_of(el2)[2:]
                    d["t"] = tier[f.stem]["tier"]
        except Exception as e:  # noqa: BLE001
            print("  건너뜀", f.stem, repr(e)[:80])
        if k % 1000 == 0:
            print(k, len(seen), flush=True)
    # 유형마다 여러 제안서에서 쓴 것부터
    out = []
    for code, name, desc in KINDS:
        cands = sorted([(s, d) for s, d in seen.items() if s[0] == code], key=lambda sd: (-len(sd[1]["docs"]), -sd[1]["n"]))
        picked = []
        for s, d in cands:
            if len(picked) >= PER:
                break
            if len(d["docs"]) < 2:
                continue
            iid = "i" + hashlib.md5(d["xml"].encode()).hexdigest()[:10]
            picked.append({"id": iid, "uses": len(d["docs"]), "w": d["wh"][0], "h": d["wh"][1], "xml": d["xml"]})
        out.append({"code": code, "name": name, "desc": desc, "items": picked})
        print(code, name, "후보", len(cands), "· 고름", len(picked))
    if not no_render:
        render(out)
    json.dump(out, open(PANE / "items.json", "w", encoding="utf-8"), ensure_ascii=False)
    print("items.json", sum(len(k["items"]) for k in out), "개")


def render(kinds):
    """Item 하나씩 빈 장에 놓고 PowerPoint 로 그림 → pane/public/pane/items/<id>.png"""
    import subprocess
    o = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"], capture_output=True, text=True, encoding="mbcs", errors="ignore").stdout or ""
    if "POWERPNT" in o:
        sys.exit("PowerPoint 가 열려 있음 — 닫고 다시")
    from pptx import Presentation
    from pptx.util import Emu
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    W, H = 12192000, 6858000
    order = []
    for k in kinds:
        for it in k["items"]:
            s = prs.slides.add_slide(prs.slide_layouts[6])
            tree = s.shapes._spTree
            el = etree.fromstring(it["xml"])
            # 가운데로
            x = el.find(f"{P}spPr/{A}xfrm") if el.tag == P + "sp" else el.find(f"{P}grpSpPr/{A}xfrm")
            w, h = it["w"], it["h"]
            x.find(A + "off").set("x", str((W - w) // 2))
            x.find(A + "off").set("y", str((H - h) // 2))
            for i, c in enumerate(el.iter()):
                if c.tag.endswith("}cNvPr"):
                    c.set("id", str(100 + i))
            tree.append(el)
            order.append((it, w, h))
    tmp = OUT / "items_src"
    tmp.mkdir(exist_ok=True)
    src = tmp / "items.pptx"
    prs.save(src)
    import pythoncom
    import win32com.client as w32
    pythoncom.CoInitialize()
    app = w32.DispatchEx("PowerPoint.Application")
    (PANE / "items").mkdir(exist_ok=True)
    try:
        pp = app.Presentations.Open(str(src), True, False, False)
        for n, (it, w, h) in enumerate(order, 1):
            png = tmp / f"{it['id']}.png"
            pp.Slides(n).Export(str(png), "PNG", 1600, 900)
            im = Image.open(png).convert("RGB")
            sx = 1600 / W
            pad = 24
            box = (max(0, int((W - w) / 2 * sx) - pad), max(0, int((H - h) / 2 * sx) - pad), min(1600, int((W + w) / 2 * sx) + pad), min(900, int((H + h) / 2 * sx) + pad))
            im = im.crop(box)
            im.thumbnail((320, 240))
            bg = Image.new("RGB", (320, 240), "white")
            bg.paste(im, ((320 - im.width) // 2, (240 - im.height) // 2))
            bg.save(PANE / "items" / f"{it['id']}.png", optimize=True)
        pp.Close()
    finally:
        app.Quit()
    # 확인용
    rows = [k for k in kinds if k["items"]]
    sheet = Image.new("RGB", (PER * 168 + 8, len(rows) * 150 + 8), "white")
    dr = ImageDraw.Draw(sheet)
    font = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 12)
    for r_, k in enumerate(rows):
        for c_, it in enumerate(k["items"]):
            t = Image.open(PANE / "items" / f"{it['id']}.png")
            t.thumbnail((160, 120))
            sheet.paste(t, (8 + c_ * 168, 8 + r_ * 150 + 18))
        dr.text((8, 8 + r_ * 150), f"{k['code']} {k['name']}", fill="black", font=font)
    sheet.save(APP / "store" / "items_sheet.jpg", quality=80)


if __name__ == "__main__":
    main()
