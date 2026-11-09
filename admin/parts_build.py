# 부품 — 구조별 「모양만」. HCG 템플릿 4개 + 창고 장표(많이 쓴 모양)에서 구조마다 서로 다른 모양을 고르고,
# 글은 모두 /// 로, 그림은 회색 「그림」 상자로 바꿔(내용 · 고객 흔적 없음) p#####.pptx 로 씀. 미리보기는 PowerPoint 로 다시 그림.
# 실행: ..\..\.venv\Scripts\python -X utf8 parts_build.py [--no-render]
#   → store/out/{pptx,img,thumb}/p#####.* + pane/public/pane/parts.json + store/parts_sheet.jpg(확인용)
# 주의: 그림 뽑기는 PowerPoint COM — 사용자가 연 PowerPoint 가 있으면 멈춤(같은 프로세스에 붙어 창을 닫을 수 있음)
import json
import re
import subprocess
import sys
import time
import zipfile

from lxml import etree
from PIL import Image, ImageDraw, ImageFont

import orphans
import regroup as R
import scrub
import store_build as B

PARTS = [("L1", "좌우 2단"), ("L2", "3~4단 기둥"), ("L3", "카드 격자"), ("L4", "위 요약 + 아래 상세"), ("L5", "좌우 비교(전→후)"), ("L6", "줄 구조"),
         ("F1", "가로 단계"), ("F2", "세로 단계"), ("F3", "순환"), ("F4", "로드맵 · 타임라인"), ("F5", "간트"),
         ("H1", "조직도"), ("H2", "로직 트리"), ("H3", "피라미드"), ("H4", "하우스"), ("H5", "프레임워크 블록"), ("H6", "중심-방사"), ("H7", "벤 · 동심원"),
         ("C1", "2×2 매트릭스"), ("C2", "옵션 비교표"), ("C3", "평가표"), ("D1", "차트 + 시사점"), ("D2", "숫자 카드"), ("T1", "표(일반)"), ("T3", "역할표")]
DOC_PREF = {"80c809c6e3": 0, "cc8b7985fa": 1, "41820cb8c8": 2, "1ead053e91": 3}  # 표준 템플릿 > Time Saver > Graphic > Ghost Pack
MAX = 24          # 구조마다 후보 수(그림 확인 뒤 parts_finalize.py 가 20개로 줄임)
MAX_TPL = 4       # 그 가운데 템플릿에서
MIN_DIFF = 40     # 고른 모양끼리 dHash(256비트) 거리 — 이보다 가까우면 같은 모양으로 봄
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
A, P = "{%s}" % NS["a"], "{%s}" % NS["p"]
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
BLANK = "///"


# ---------- 모양만 남기기 ----------
def _blank_paras(root, keep_text=True):
    n = 0
    for p in root.iter(A + "p"):
        runs = [c for c in p if c.tag in (A + "r", A + "fld", A + "br")]
        if not runs or not "".join(t.text or "" for t in p.iter(A + "t")).strip():
            for c in runs:  # 빈 칸 · 줄바꿈만 있는 문단 — 글자 흔적만 지움
                for t in c.iter(A + "t"):
                    t.text = ""
            continue
        first = next((c for c in runs if c.tag in (A + "r", A + "fld")), None)
        for c in runs:
            if c is not first:
                p.remove(c)
        if first is not None and first.tag == A + "fld":  # 필드(쪽 번호 등)는 보통 글로
            first.tag = A + "r"
            for k in list(first.attrib):
                del first.attrib[k]
            for c in list(first):
                if c.tag not in (A + "rPr", A + "t"):
                    first.remove(c)
        if first is not None:
            t = first.find(A + "t")
            if t is None:
                t = etree.SubElement(first, A + "t")
            t.text = BLANK if keep_text else ""
            n += 1
    for el in root.iter():  # 그림 설명 · 도형 제목(대체 글)
        if el.tag.endswith("}cNvPr"):
            for att in ("descr", "title"):
                el.attrib.pop(att, None)
    return n


def _pic_box(pic, sw):
    """그림 → 같은 자리 회색 상자. 아주 작은 것은 글 없이."""
    xfrm = pic.find(f"{P}spPr/{A}xfrm")
    return _box(xfrm, pic.find(f".//{P}cNvPr"), sw, "그림")


def _box(xfrm, cnv, sw, label):
    """같은 자리 회색 상자. xfrm 은 a:xfrm(그림) 또는 p:xfrm(개체 틀) — a:xfrm 으로 옮김. 아주 작은 것은 글 없이."""
    sp = etree.Element(P + "sp", nsmap={"a": NS["a"], "p": NS["p"]})
    nv = etree.SubElement(sp, P + "nvSpPr")
    etree.SubElement(nv, P + "cNvPr", id=cnv.get("id") if cnv is not None else "9999", name=f"{label} 자리")
    etree.SubElement(nv, P + "cNvSpPr")
    etree.SubElement(nv, P + "nvPr")
    spPr = etree.SubElement(sp, P + "spPr")
    w = 0
    if xfrm is not None:
        ax = etree.SubElement(spPr, A + "xfrm", {k: v for k, v in xfrm.attrib.items()})
        for c in xfrm:
            if etree.QName(c).localname in ("off", "ext"):
                ax.append(etree.fromstring(etree.tostring(c)))
        ext = xfrm.find(A + "ext")
        w = int(ext.get("cx", "0")) if ext is not None else 0
    g = etree.SubElement(spPr, A + "prstGeom", prst="rect")
    etree.SubElement(g, A + "avLst")
    f = etree.SubElement(spPr, A + "solidFill")
    etree.SubElement(f, A + "srgbClr", val="D9D9D9")
    ln = etree.SubElement(spPr, A + "ln")
    etree.SubElement(ln, A + "noFill")
    tb = etree.SubElement(sp, P + "txBody")
    etree.SubElement(tb, A + "bodyPr", anchor="ctr", lIns="0", rIns="0")
    etree.SubElement(tb, A + "lstStyle")
    p = etree.SubElement(tb, A + "p")
    etree.SubElement(p, A + "pPr", algn="ctr")
    if w > sw * 0.06:
        r = etree.SubElement(p, A + "r")
        rp = etree.SubElement(r, A + "rPr", lang="ko-KR", sz="1000")
        sf = etree.SubElement(rp, A + "solidFill")
        etree.SubElement(sf, A + "srgbClr", val="7F7F7F")
        etree.SubElement(r, A + "t").text = label
    return sp


def _strip_media(root, sw):
    """장 안 그림 · 붙인 개체(OLE) · 그림 채우기를 회색으로. 바꾼 수."""
    n = 0
    MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
    # 붙인 개체(엑셀 · think-cell 등) — 개체 틀 통째로(대체 그림 p:pic 이 안에 있어 따로 바꾸면 파일이 깨짐)
    for ole in list(root.iter(P + "oleObj")):
        gf = ole
        while gf is not None and gf.tag != P + "graphicFrame":
            gf = gf.getparent()
        if gf is None or gf.getparent() is None:
            continue
        top = gf
        if gf.getparent().tag in (MC + "Choice", MC + "Fallback"):
            top = gf.getparent().getparent()
        if top.getparent() is None:
            continue
        top.getparent().replace(top, _box(gf.find(P + "xfrm"), gf.find(f"{P}nvGraphicFramePr/{P}cNvPr"), sw, "개체"))
        n += 1
    for pic in list(root.iter(P + "pic")):
        if pic.getparent() is not None:
            pic.getparent().replace(pic, _pic_box(pic, sw))
            n += 1
    # 도형 그림 채우기 · 장 배경 그림
    for bf in list(root.iter(A + "blipFill")):
        f = etree.Element(A + "solidFill")
        etree.SubElement(f, A + "srgbClr", val="D9D9D9")
        bf.getparent().replace(bf, f)
        n += 1
    return n


def skeleton(src, dst, keep_charts):
    zin = zipfile.ZipFile(src)
    pres = etree.fromstring(zin.read("ppt/presentation.xml"))
    sw = int(pres.find(f"{P}sldSz").get("cx"))
    out = {}
    for n in zin.namelist():
        out[n] = zin.read(n)
    blanks = 0
    for n in list(out):
        if re.match(r"ppt/slides/slide\d+\.xml$", n):
            root = etree.fromstring(out[n])
            _strip_media(root, sw)
            blanks += _blank_paras(root)
            out[n] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        elif re.match(r"ppt/diagrams/(data|drawing)\d+\.xml$", n):
            root = etree.fromstring(out[n])
            blanks += _blank_paras(root)
            out[n] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        elif re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n):
            root = etree.fromstring(out[n])
            _blank_paras(root, keep_text=False)
            out[n] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # 장 관계: 지운 그림 · 개체 · 메모(댓글) 연결 빼기(장 파일 이름은 원본 번호 그대로 — slide39.xml 등)
    for sn in [n for n in out if re.match(r"ppt/slides/slide\d+\.xml$", n)]:
        rel_n = sn.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        if rel_n not in out:
            continue
        slide = out[sn].decode("utf-8")
        rels = etree.fromstring(out[rel_n])
        for rel in list(rels):
            typ, rid = rel.get("Type", ""), rel.get("Id")
            used = re.search(r'r:[A-Za-z]+="%s"' % re.escape(rid), slide)
            if typ.endswith("/comments") or (not used and re.search(r"/(image|oleObject|package|chart|vmlDrawing|audio|video|media)$", typ)):
                rels.remove(rel)
        out[rel_n] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    # 어디서도 가리키지 않는 그림 · 메모 파일 빼기
    targets = set()
    for n, data in out.items():
        if not n.endswith(".rels"):
            continue
        base = n.replace("_rels/", "").rsplit(".rels", 1)[0]
        folder = base.rsplit("/", 1)[0] if "/" in base else ""
        for rel in etree.fromstring(data):
            if rel.get("TargetMode") == "External":
                continue
            t = rel.get("Target", "")
            parts = (t.lstrip("/") if t.startswith("/") else f"{folder}/{t}").split("/")
            norm = []
            for x in parts:
                if x == "..":
                    if norm:
                        norm.pop()
                elif x and x != ".":
                    norm.append(x)
            targets.add("/".join(norm))
    for n in list(out):
        if re.match(r"ppt/(media|embeddings|comments|drawings)/|ppt/comments", n) and n not in targets:
            del out[n]
    orphans.clean(out)  # 숨은 다른 장 · viewProps 빼고 형식 등록도 정리
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", out.pop("[Content_Types].xml"))
        for n, data in out.items():
            z.writestr(n, data)
    return blanks


def box_ratio(path):
    """회색 그림 · 개체 상자가 장에서 차지하는 비율 — 크면 모양이 거의 안 남음."""
    z = zipfile.ZipFile(path)
    sz = etree.fromstring(z.read("ppt/presentation.xml")).find(f"{P}sldSz")
    area = int(sz.get("cx")) * int(sz.get("cy"))
    tot = 0
    for n in z.namelist():
        if re.match(r"ppt/slides/slide\d+\.xml$", n):
            for sp in etree.fromstring(z.read(n)).iter(P + "sp"):
                c = sp.find(f"{P}nvSpPr/{P}cNvPr")
                ext = sp.find(f"{P}spPr/{A}xfrm/{A}ext")
                if c is not None and c.get("name") in ("그림 자리", "개체 자리") and ext is not None and sp.getparent().tag == P + "spTree":
                    tot += int(ext.get("cx", 0)) * int(ext.get("cy", 0))
    return tot / area


def leftover(path):
    """모양만 남긴 파일에 /// · 그림 말고 글이 남았는지(쪽 번호 숫자는 봐줌)."""
    z = zipfile.ZipFile(path)
    bad = []
    for n in z.namelist():
        if re.match(r"ppt/(slides/slide|diagrams/(data|drawing)|notesSlides/notesSlide)\d+\.xml$", n):
            for t in etree.fromstring(z.read(n)).iter(A + "t"):
                s = (t.text or "").strip()
                if s and s not in (BLANK, "그림") and not s.isdigit():
                    bad.append((n, s[:30]))
    return bad


# ---------- PowerPoint 로 그림 ----------
def ppt_pids():
    o = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True, encoding="mbcs", errors="ignore").stdout or ""
    return {line.split(",")[1].strip('"') for line in o.splitlines() if "POWERPNT" in line}


def render(jobs):
    if ppt_pids():
        sys.exit("PowerPoint 가 열려 있음 — 저장하고 닫은 뒤 다시 실행(그림 뽑기가 같은 PowerPoint 에 붙어 창을 닫을 수 있음)")
    import pythoncom
    import win32com.client as w32
    pythoncom.CoInitialize()
    app = w32.DispatchEx("PowerPoint.Application")
    fails = []
    try:
        for n, (src, png) in enumerate(jobs, 1):
            if png.exists() and png.stat().st_mtime > src.stat().st_mtime:
                continue  # 이미 그림(다시 돌릴 때 건너뜀)
            try:
                p = app.Presentations.Open(str(src), True, False, False)
                w, h = p.PageSetup.SlideWidth, p.PageSetup.SlideHeight
                p.Slides(1).Export(str(png), "PNG", 1600, round(1600 * h / w))
                p.Close()
            except Exception as e:  # noqa: BLE001 — PowerPoint 가 못 여는 파일은 부품에서 뺌
                fails.append(src.stem)
                print("  못 엶", src.stem, str(e)[:80], flush=True)
            if n % 25 == 0:
                print("그림", n, "/", len(jobs), flush=True)
        return fails
    finally:
        try:
            app.Quit()
        except Exception as e:  # noqa: BLE001
            print("quit", e)


# ---------- 고르기 ----------
def main():
    no_render = "--no-render" in sys.argv
    idx, conv, refs = B.load_index()
    names = B.known_names()
    rows = [json.loads(line) for line in open(B.WORK / "ref" / "review" / "tags_all.jsonl", encoding="utf-8")]
    slides = {}
    for line in open(B.WORK / "ref" / "review" / "slides_all.jsonl", encoding="utf-8"):
        d = json.loads(line)
        slides[d["k"]] = d
    meta = json.load(open(B.OUT / "meta.json", encoding="utf-8"))
    guide = re.compile(r"SVG|PNG|JPEG|아이콘의|색상 적용법|사용법|활용법|가이드|Guide")
    tmp = B.OUT / "parts_src"
    tmp.mkdir(exist_ok=True)
    out, jobs, picked_all = [], [], []

    def has(path, pat):
        try:
            return any(re.match(pat, n) for n in zipfile.ZipFile(path).namelist())
        except zipfile.BadZipFile:
            return True

    for code, name in PARTS:
        chosen = []  # (k, src_pptx, thumb_png, origin, uses)

        def far(th):
            h = R.dhash(th)
            if h is None:
                return None
            for c in chosen:
                if c[5] is not None and bin(h ^ c[5]).count("1") < MIN_DIFF:
                    return None
            return h

        # 1) 템플릿 — 이미 일반 글(Text Here 등)이지만 같은 방식으로 지움
        tpl = [r for r in rows if r["rep"] == r["k"] and r["doc"] in DOC_PREF and r["s"] == code
               and not guide.search((r.get("title") or "") + " " + (slides.get(r["k"], {}).get("text") or ""))
               and len(slides.get(r["k"], {}).get("text") or "") >= 20]
        tpl.sort(key=lambda r: (B_GRADE.get(r["re"], 3), DOC_PREF[r["doc"]], r["k"]))
        for r in tpl:
            if sum(1 for c in chosen if c[3] == "템플릿") >= MAX_TPL:
                break
            pp = B.OUT / "pptx" / f"s{r['k']:05d}.pptx"
            if not pp.exists():
                src = B.source_path(r, idx, conv, refs)
                if not src:
                    continue
                try:
                    B.cut_slide(src, r["i"], pp)
                    scrub.scrub_file(pp, names)
                except Exception as e:  # noqa: BLE001
                    print("  자르기 실패", r["k"], e)
                    continue
            h = far(r["png"])
            if h is None:
                continue
            chosen.append((r["k"], pp, r["png"], "템플릿", 0, h))
        # 2) 창고 — 여러 제안서에서 쓴 모양부터(묶음 안 모양 대표만). 차트 · 붙인 개체 · 그림 많은 장은 뺌
        wh = [m for m in meta if m["s"] == code and m.get("var_rep") == 1 and 40 <= len(m.get("text", "")) <= 1600]
        wh.sort(key=lambda m: (-m.get("var_n", 1), -(m["tier"] == 1), -m.get("fam_docs", 1), -bool(m["star"]), m["k"]))
        for m in wh:
            if len(chosen) >= MAX + 4:  # 지운 뒤 빠질 것(그림이 대부분인 장)을 생각해 넉넉히
                break
            if (slides.get(m["k"], {}).get("n") or {}).get("pic", 0) > 3:
                continue
            pp = B.OUT / "pptx" / f"{m['id']}.pptx"
            if not pp.exists() or has(pp, r"ppt/(charts|embeddings)/"):
                continue
            th = B.OUT / "thumb" / f"{m['id']}.jpg"
            h = far(th)
            if h is None:
                continue
            chosen.append((m["k"], pp, th, "지난 제안서", m.get("var_n", 1), h))
        vars_ = []
        for k, pp, th, origin, uses, _h in chosen:
            if len(vars_) >= MAX:
                break
            pid = f"p{k:05d}"
            dst = B.OUT / "pptx" / f"{pid}.pptx"
            skeleton(pp, dst, keep_charts=(origin == "템플릿"))
            bad = leftover(dst)
            if bad:
                print("  글 남음", pid, bad[:3])
                dst.unlink()
                continue
            if box_ratio(dst) > 0.3:  # 그림이 대부분인 장 — 모양이 거의 안 남아 뺌
                dst.unlink()
                continue
            jobs.append((dst, tmp / f"{pid}.png"))
            vars_.append({"id": pid, "title": f"{name} {len(vars_) + 1}", "origin": origin, "uses": uses})
            picked_all.append((code, name, pid))
        # 창고 · 템플릿에 이 구조로 분류된 장 수
        cnt = sum(1 for m in meta if m["s"] == code) + sum(1 for r in rows if r["rep"] == r["k"] and r["doc"] in DOC_PREF and r["s"] == code)
        if vars_:
            out.append({"code": code, "name": name, "struct": B.STRUCT.get(code, code), "count": cnt, "variants": vars_})
        print(code, name, f"모양 {len(vars_)} (템플릿 {sum(v['origin'] == '템플릿' for v in vars_)}) · 분류된 장 {cnt}", flush=True)

    if not no_render:
        t0 = time.time()
        fails = set(render(jobs))
        print(f"그림 {len(jobs)}장 · 못 엶 {len(fails)} · {time.time() - t0:.0f}s")
        for p in out:
            p["variants"] = [v for v in p["variants"] if v["id"] not in fails]
        out = [p for p in out if p["variants"]]
    for dst, png in jobs:
        if not png.exists():
            continue
        im = Image.open(png).convert("RGB")
        pid = dst.stem
        for d, w, q in (("img", 960, 84), ("thumb", 360, 80)):
            im.resize((w, round(im.height * w / im.width)), Image.LANCZOS).save(B.OUT / d / f"{pid}.jpg", quality=q, optimize=True)
    # 고르지 않은 부품 파일 지움(올리지 않게)
    keep = {v["id"] for p in out for v in p["variants"]}
    for d, ext in (("pptx", "pptx"), ("img", "jpg"), ("thumb", "jpg")):
        for f in (B.OUT / d).glob(f"p*.{ext}"):
            if f.stem not in keep:
                f.unlink()
    dst = B.APP / "pane" / "public" / "pane" / "parts.json"
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    # 확인용 그림(사람이 한 번 봄)
    W, H, C = 240, 135, 10
    img = Image.new("RGB", (C * (W + 6) + 6, len(out) * (H + 22) + 6), "white")
    dr = ImageDraw.Draw(img)
    font = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 12)
    for i, p in enumerate(out):
        for j, v in enumerate(p["variants"]):
            f = B.OUT / "thumb" / f"{v['id']}.jpg"
            if not f.exists():
                continue
            t = Image.open(f)
            t.thumbnail((W, H))
            x, y = 6 + j * (W + 6), 6 + i * (H + 22)
            img.paste(t, (x, y + 18))
            dr.text((x, y + 2), f"{p['code']} {p['name']} · {v['id']} · {v['origin']}", fill="black", font=font)
    img.save(B.APP / "store" / "parts_sheet.jpg", quality=80)
    print("parts.json", len(out), "부품 ·", sum(len(p["variants"]) for p in out), "모양")


B_GRADE = {"상": 0, "중": 1, "中": 1, "하": 2}

if __name__ == "__main__":
    main()
