# Draft 마무리(판 3): 넘침 → 자동 줄바꿈 · 줄여 쓰기 요청문 / 그림 속 글 딱지 / 파일 속성 새로 쓰기
import collections, json, os, re, sys, zipfile, hashlib
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

D = os.path.dirname(os.path.abspath(__file__))  # 작업 폴더(run.py 가 바꿈)
HERE = D
sys.path.insert(0, D)
CLIENT = "TCK"
CLIENT_NAMES = ["TCK", "티씨케이"]
TITLE = "평가제도 고도화"
INPUT = ""  # 고객 내용 파일(hcgd.py 가 바꿈)
META = ""  # 창고 목록 meta.json(hcgd.py 가 바꿈)


def client_info():
    t = open(INPUT, encoding="utf-8").read()
    return t.split("\n", 1)[1] if t.startswith("/hcg") else t
from draft import ptext, set_para, plain, tidy, KEEP_ROLES  # noqa: E402

def stype(sh):
    """도형 종류. python-pptx 가 모르는 도형(특수 자동 도형)은 None."""
    try:
        return sh.shape_type
    except NotImplementedError:
        return None


def keep_slides():
    """그대로 넣은 장(HCG 소개 등) 번호 — 고쳐 쓰기 · 이름 검사 · 서식 손질에서 뺌."""
    try:
        order = json.load(open(os.path.join(D, "order.json"), encoding="utf-8"))
    except (OSError, ValueError):
        return set()
    return {i for i, o in enumerate(order, 1) if o.get("role") in KEEP_ROLES}


def load(p):
    return json.load(open(p, encoding="utf-8-sig")) if os.path.exists(p) else []


def as_list(x):
    return x if isinstance(x, list) else ([x] if x else [])


def walk_abs(shapes, tf=None):
    """도형을 장 기준 좌표로(묶음 안 도형 포함)."""
    for sh in shapes:
        if stype(sh) == MSO_SHAPE_TYPE.GROUP:
            x = sh._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
            off, ext = x.find(qn("a:off")), x.find(qn("a:ext"))
            choff, chext = x.find(qn("a:chOff")), x.find(qn("a:chExt"))
            sx = int(ext.get("cx")) / max(1, int(chext.get("cx")))
            sy = int(ext.get("cy")) / max(1, int(chext.get("cy")))
            g = (int(off.get("x")), int(off.get("y")), int(choff.get("x")), int(choff.get("y")), sx, sy)
            inner = lambda a, g=g: (g[0] + (a[0] - g[2]) * g[4], g[1] + (a[1] - g[3]) * g[5], a[2] * g[4], a[3] * g[5])
            comp = (lambda a, f=tf, h=inner: f(h(a))) if tf else inner
            yield from walk_abs(sh.shapes, comp)
        else:
            box = (sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0)
            yield sh, (tf(box) if tf else box)


def overflow_new(over, base):
    """원본에서도 넘치던 도형(디자인상 넘침)은 빼고 새로 넘친 것만."""
    b = {(x["slide"], x["id"]): x for x in as_list(base)}
    out = []
    for x in as_list(over):
        if x["bh"] <= x["h"] + 6 and x["bw"] <= x["w"] + 6:  # 6pt 이하 = 여백 차이, 눈에 안 띔
            continue
        o = b.get((x["slide"], x["id"]))
        if o is None or x["bh"] > o["bh"] * 1.1 + 2 or x["bw"] > o["bw"] * 1.1 + 2:
            out.append(x)
    return out


def fixwrap(deck, over_p, base_p, out):
    """한 줄로 쓰는 도형(줄바꿈 끔)에서 글이 옆으로 넘치면 도형 너비에 맞춰 자동 줄바꿈을 켬."""
    prs = Presentation(deck)
    new = overflow_new(load(over_p), load(base_p))
    want = {(x["slide"], x["id"]) for x in new if x["bw"] > x["w"] + 4}
    n = 0
    for i, s in enumerate(prs.slides, 1):
        for sh, _ in walk_abs(s.shapes):
            if (i, sh.shape_id) in want and sh.has_text_frame:
                bp = sh.text_frame._txBody.find(qn("a:bodyPr"))
                if bp.get("wrap") == "none":
                    bp.set("wrap", "square"); n += 1
    prs.save(out)
    print("자동 줄바꿈 켬", n)


def fixprompt(deck, over_p, base_p, outp):
    prs = Presentation(deck)
    new = overflow_new(load(over_p), load(base_p))
    tck = client_info()
    rows = []
    want = {(x["slide"], x["id"]): x for x in new}
    for i, s in enumerate(prs.slides, 1):
        for sh, _ in walk_abs(s.shapes):
            x = want.get((i, sh.shape_id))
            if not x or not sh.has_text_frame:
                continue
            ratio = min(x["h"] / max(1, x["bh"]), 1.0)
            if x["bw"] > x["w"] + 4:
                ratio = min(ratio, x["w"] / max(1, x["bw"]))
            for k, para in enumerate(sh.text_frame.paragraphs):
                t = ptext(para)
                if t:
                    lim = max(4, int(len(plain(t)) * ratio * 0.9))
                    rows.append(f"{i}:{sh.shape_id}:{k} | 최대 {lim}자 | 줄 {t.count(chr(10)) + 1} | {json.dumps(t, ensure_ascii=False)}")
    p = ("너는 HCG 제안서 Draft 를 다듬는 컨설턴트다. 아래 항목들은 도형 안에서 글이 넘쳤다. 뜻은 살리고 「최대 N자」 안으로 줄여 쓴다(띄어쓰기 포함, ** 는 세지 않음). "
         "「줄 N」은 줄 수 — 줄바꿈(\\n) 자리와 줄 수를 그대로 지킨다. **…** 는 굵게 표시 — 같은 성격의 구절에 그대로 둔다. 고객 정보는 참고만.\n\n# 고객 정보\n" + tck
         + "\n\n# 줄일 항목 (id | 최대 글자 | 줄 수 | 지금 글)\n" + "\n".join(rows)
         + '\n\n# 출력\nJSON 하나만: {"texts": {"id": "줄인 글", ...}}')
    open(outp, "w", encoding="utf-8").write(p)
    print("줄일 항목", len(rows), "도형", len(want))


def apply_texts(deck, ans_p, out):
    prs = Presentation(deck)
    raw = json.load(open(ans_p, encoding="utf-8"))
    t = raw["result"] if "result" in raw else raw
    if isinstance(t, str):
        t = json.loads(re.sub(r"^```(json)?|```$", "", t.strip(), flags=re.M))
    texts = t["texts"]
    n = 0
    for key, para in iter_paras(prs):
        v = texts.get(key)
        if v is not None and set_para(para, v):
            n += 1
    tidy(prs, keep_slides())
    prs.save(out)
    print("줄여 씀", n)


def pics(deck, pngdir, cropdir):
    """넓이 2% 이상 그림을 장 그림(1920px)에서 잘라 OCR 할 준비."""
    from PIL import Image
    prs = Presentation(deck)
    SW, SH = prs.slide_width, prs.slide_height
    os.makedirs(cropdir, exist_ok=True)
    rows = []
    for i, s in enumerate(prs.slides, 1):
        img = None
        for sh, (l, t, w, h) in walk_abs(s.shapes):
            if stype(sh) != MSO_SHAPE_TYPE.PICTURE or w * h / (SW * SH) < 0.02:
                continue
            if img is None:
                img = Image.open(os.path.join(pngdir, f"R_{i:02d}.png"))
            X, Y = img.size
            box = (max(0, int(l / SW * X)), max(0, int(t / SH * Y)), min(X, int((l + w) / SW * X)), min(Y, int((t + h) / SH * Y)))
            if box[2] - box[0] < 20 or box[3] - box[1] < 20:
                continue
            key = f"s{i:02d}_{sh.shape_id}"
            c = img.crop(box)
            if c.width < 600:  # 작은 그림은 키워서 OCR
                c = c.resize((c.width * 2, c.height * 2))
            c.save(os.path.join(cropdir, key + ".png"))
            rows.append({"key": key, "slide": i, "id": sh.shape_id, "box": [int(l), int(t), int(w), int(h)], "area": round(w * h / (SW * SH), 3)})
    json.dump(rows, open(os.path.join(D, "pics.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print("그림", len(rows))


TAG_FILL, TAG_LINE, TAG_TEXT = RGBColor(0xFF, 0xD8, 0x00), RGBColor(0xC0, 0x00, 0x00), RGBColor(0x1A, 0x1A, 0x1A)
HOLE = re.compile(r"○{2,}|◯{2,}|OO(?![A-Za-z])")


def tag_box(label):
    """딱지 크기(pt): 12pt 굵게, 한글 12pt · 그 밖 7pt 폭으로 어림."""
    w = sum(12.5 if re.match(r"[가-힣ㄱ-ㅎ○◯]", ch) else 7 for ch in label) + 18
    return Pt(max(96, min(w, 330))), Pt(24)


def stick(deck, ocr_p, out, extra_p=""):
    """확인할 곳마다 노란 딱지(장 위, 이름은 「HCG확인」으로 시작해 작업창에서 한 번에 찾아 지움).
    붙이는 곳: 글이 든 그림 · 다른 고객 그림(그림은 지우지 않음) · ○○ 자리 · Claude 확인 항목 · 넘침 · 겹침."""
    prs = Presentation(deck)
    SW, SH = prs.slide_width, prs.slide_height
    skip = keep_slides()
    ocr = load(ocr_p)
    ocr = ocr if isinstance(ocr, dict) else {}
    rows = load(os.path.join(D, "pics.json"))
    midx = load(os.path.join(HERE, "media_index.json"))
    checks = [c for c in load(os.path.join(D, "checks.json")) if c.get("kind") not in ("그림", "값")]  # 다시 돌려도 겹치지 않게
    marks = {}  # (장, 도형 id) → [딱지 말]

    def mark(i, sid, label):
        if i in skip:
            return
        lst = marks.setdefault((i, int(sid)), [])
        if label not in lst:
            lst.append(label)

    n = {"그림 속 글": 0, "다른 고객 그림": 0, "값 넣기": 0, "Claude": 0, "넘침 · 겹침": 0}
    for r in rows:
        txt = re.sub(r"\s", "", (ocr.get(r["key"]) or {}).get("text") or "")
        if len(re.findall(r"[가-힣A-Za-z0-9]", txt)) >= 8:
            mark(r["slide"], r["id"], "그림 속 글"); n["그림 속 글"] += 1
            checks.append({"slide": r["slide"], "note": f"그림 속 글(「{txt[:24]}…」) — 글을 못 바꿈, 그림 바꾸기 · 지우기", "kind": "그림"})
    for i, s in enumerate(prs.slides, 1):
        if i in skip:
            continue
        for sh, box in walk_abs(s.shapes):
            if stype(sh) == MSO_SHAPE_TYPE.PICTURE:
                try:
                    e = midx.get(hashlib.md5(sh.image.blob).hexdigest())
                except Exception:
                    e = None
                cl = [x for x in (e or {}).get("clients", []) if x]
                l_, t_, w_, h_ = box
                icon = w_ * h_ < SW * SH * 0.015 and 0.55 <= w_ / max(h_, 1) <= 1.8  # 작은 정사각 아이콘(글 없음)은 다른 고객 것이어도 그대로 씀
                if e and len(cl) == 1 and e["ndocs"] <= 2 and not icon:
                    mark(i, sh.shape_id, "다른 고객 그림"); n["다른 고객 그림"] += 1
                    checks.append({"slide": i, "note": f"다른 고객({cl[0]}) 자료 그림 — 바꾸거나 지우기", "kind": "그림"})
            txt = ""
            if sh.has_text_frame:
                txt = sh.text_frame.text
            elif getattr(sh, "has_table", False) and sh.has_table:
                txt = " ".join(c.text for row in sh.table.rows for c in row.cells)
            if txt and HOLE.search(txt) and not txt.startswith("확인 ·"):
                mark(i, sh.shape_id, "값 넣기"); n["값 넣기"] += 1
                checks.append({"slide": i, "note": f"○○ 자리 — 값 넣기(「{re.sub(chr(10), ' ', txt)[:20]}…」)", "kind": "값"})
    for c in checks:
        cid = str(c.get("id") or "")
        m = re.match(r"(\d+):(\d+)", cid)
        if c.get("kind") in ("그림", "값") or not m:
            continue
        mark(int(m.group(1)), m.group(2), c.get("tag") or "내용 확인"); n["Claude"] += 1
    for x in load(extra_p) if extra_p else []:
        mark(x["slide"], x["id"], x["tag"]); n["넘침 · 겹침"] += 1
    # 자리 없는 Claude 확인 → 장 오른쪽 위에 하나로
    loose = collections.defaultdict(list)
    for c in checks:
        if c.get("kind") in (None, "Claude", "차트") and not re.match(r"\d+:\d+", str(c.get("id") or "")) and c.get("slide") and c["slide"] not in skip:
            loose[c["slide"]].append(c.get("tag") or "내용 확인")
    # 한 장에 같은 종류 딱지(값 넣기 · 다른 고객 그림 · 그림 속 글)가 4개 넘으면 하나로 묶어 장 오른쪽 위에
    MERGE = ("값 넣기", "다른 고객 그림", "그림 속 글")
    per = collections.defaultdict(collections.Counter)
    for (i, sid), labels in marks.items():
        for lb in labels:
            if lb in MERGE:
                per[i][lb] += 1
    for i, cnt in per.items():
        for lb, k in cnt.items():
            if k > 3:
                for key in [key for key in marks if key[0] == i and lb in marks[key]]:
                    marks[key].remove(lb)
                loose[i].append(f"{lb} {k}곳")
    marks = {key: v for key, v in marks.items() if v}
    # 그래도 한 장에 6개 넘으면 위에서부터 6개만 도형 옆에, 나머지는 장 위 한 딱지로
    by_slide = collections.defaultdict(list)
    for key in sorted(marks):
        by_slide[key[0]].append(key)
    for i, keys in by_slide.items():
        for key in keys[6:]:
            loose[i] += marks.pop(key)
    placed = collections.defaultdict(list)
    obst = collections.defaultdict(list)  # 장 → 글이 있는 도형 자리(장 넓이 25% 넘는 바탕 도형은 빼고)
    for i, s in enumerate(prs.slides, 1):
        for sh, (l, t, w, h) in walk_abs(s.shapes):
            has = (sh.has_text_frame and sh.text_frame.text.strip()) or (getattr(sh, "has_table", False) and sh.has_table)
            if has and w * h < SW * SH * 0.25:
                obst[i].append((int(l), int(t), int(w), int(h)))

    def ov(a, b):
        return max(0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])) * max(0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))

    def cost(i, r):
        return (sum(ov(r, o) for o in obst[i]) + 3 * sum(ov(r, p) for p in placed[i])) / float(r[2] * r[3])

    def put(i, x, y, label, sid, box=None):
        """딱지 자리: 도형 위 바깥(왼 · 오른) → 아래 바깥 → 오른쪽 · 왼쪽 바깥 중 다른 글 · 딱지를 가장 덜 가리는 곳. 그래도 딱지끼리 겹치면 오른쪽 → 위로 비킴."""
        s = prs.slides[i - 1]
        bw, bh = tag_box(label)
        cands = [(x, y)]
        if box:
            l, t, w, h = box
            g = int(Pt(2))
            cands = [(l, t - bh), (l + w - bw, t - bh), (l, t + h + g), (l + w - bw, t + h + g), (l + w + g, t), (l - bw - g, t), (l, t + g)]
        best = None
        for cx, cy in cands:
            cx, cy = int(min(max(0, cx), SW - bw)), int(min(max(0, cy), SH - bh))
            c = cost(i, (cx, cy, bw, bh))
            if best is None or c < best[0] - 1e-6:
                best = (c, cx, cy)
            if c == 0:
                break
        _, x, y = best
        x0 = x
        for _ in range(30):
            hit = [(a, b, w, h) for a, b, w, h in placed[i] if x < a + w and a < x + bw and y < b + h and b < y + bh]
            if not hit:
                break
            nx = max(a + w for a, b, w, h in hit) + int(Pt(4))
            if nx + bw <= SW:
                x = nx
            else:
                x, y = x0, int(max(0, y - bh - Pt(2)))
        placed[i].append((x, y, bw, bh))
        tag = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, bw, bh)
        tag.name = f"HCG확인 {label} {sid}"
        tag.fill.solid(); tag.fill.fore_color.rgb = TAG_FILL
        tag.line.color.rgb = TAG_LINE; tag.line.width = Pt(1.5)
        tag.shadow.inherit = False
        tf = tag.text_frame; tf.margin_left = tf.margin_right = Pt(6); tf.margin_top = tf.margin_bottom = Pt(1)
        tf.word_wrap = False
        tf.text = label
        r = tf.paragraphs[0].runs[0]; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = TAG_TEXT; r.font.name = "맑은 고딕"

    boxes = {}
    for i, s in enumerate(prs.slides, 1):
        for sh, box in walk_abs(s.shapes):
            boxes[(i, sh.shape_id)] = box
    for (i, sid), labels in sorted(marks.items()):
        box = boxes.get((i, sid))
        if not box:
            loose[i] += labels; continue
        l, t, w, h = box
        label = "확인 · " + " / ".join(labels[:2])
        put(i, l, t, label, sid, box=box)  # 도형 바깥, 다른 글을 가리지 않는 곳
    for i, labels in loose.items():
        uq = list(dict.fromkeys(labels))
        label = "확인 · " + " / ".join(uq[:2]) + (f" 외 {len(uq) - 2}건" if len(uq) > 2 else "")
        bw, _ = tag_box(label)
        put(i, SW - bw - Pt(10), Pt(6), label, "장")
    prs.save(out)
    json.dump(checks, open(os.path.join(D, "checks.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(n, ensure_ascii=False))


def overlaps(base_p, new_p, out_p=None):
    """글이 이웃 도형을 새로 덮는 곳: 글 자리(Bound)와 다른 도형 자리가 겹치는 넓이가 원본보다 커진 경우."""
    def inter(a, b):
        w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]); h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
        return w * h if w > 0 and h > 0 else 0
    def index(rows):
        by = {}
        for r in rows:
            by.setdefault(r["slide"], {})[r["id"]] = r
        return by
    B, N = index(load(base_p)), index(load(new_p))
    found = []
    for s, shapes in N.items():
        for t in shapes.values():
            if "bw" not in t or t["t"].startswith("확인 ·"):
                continue
            tb = (t["bx"], t["by"], t["bw"], t["bh"]); trect = (t["x"], t["y"], t["w"], t["h"])
            ob_t = B.get(s, {}).get(t["id"])
            for o in shapes.values():
                if o["id"] == t["id"] or o["w"] * o["h"] > 960 * 540 * 0.5 or o["type"] in (9,) or o["t"].startswith("확인 ·"):  # 바탕 · 선 · 딱지는 뺌
                    continue
                orect = (o["x"], o["y"], o["w"], o["h"])
                if inter(orect, trect) >= trect[2] * trect[3] * 0.9:
                    continue  # 글 상자를 감싼 판
                a = inter(tb, orect)
                if a < 30:
                    continue
                before = 0
                if ob_t and "bw" in ob_t and o["id"] in B.get(s, {}):
                    oo = B[s][o["id"]]
                    before = inter((ob_t["bx"], ob_t["by"], ob_t["bw"], ob_t["bh"]), (oo["x"], oo["y"], oo["w"], oo["h"]))
                if a > before * 1.3 + 20:
                    found.append({"slide": s, "id": t["id"], "other": o["id"], "area": round(a), "before": round(before), "text": t["t"]})
    if out_p:
        json.dump(found, open(out_p, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print("새 겹침", len(found))
    for f in found:
        print(f)


def fixprompt_ovl(deck, ovl_p, outp):
    """겹침 난 글 상자를 원래 길이의 75% 안으로 줄여 쓰게 하는 요청문(줄 수 유지)."""
    prs = Presentation(deck)
    want = {(x["slide"], x["id"]) for x in load(ovl_p) if x["area"] >= 60}
    tck = client_info()
    rows = []
    for i, s in enumerate(prs.slides, 1):
        for sh, _ in walk_abs(s.shapes):
            if (i, sh.shape_id) in want and sh.has_text_frame:
                for k, para in enumerate(sh.text_frame.paragraphs):
                    t = ptext(para)
                    if t:
                        rows.append(f"{i}:{sh.shape_id}:{k} | 최대 {max(4, int(len(plain(t)) * 0.75))}자 | 줄 {t.count(chr(10)) + 1} | {json.dumps(t, ensure_ascii=False)}")
    p = ("너는 HCG 제안서 Draft 를 다듬는 컨설턴트다. 아래 항목들은 글이 옆 도형을 덮었다. 뜻은 살리고 「최대 N자」 안으로 줄여 쓴다(띄어쓰기 포함, ** 는 세지 않음). "
         "「줄 N」은 줄 수 — 줄바꿈(\\n) 자리와 줄 수를 그대로 지킨다. **…** 는 굵게 표시 — 같은 성격의 구절에 그대로 둔다. 고객 정보는 참고만.\n\n# 고객 정보\n" + tck
         + "\n\n# 줄일 항목 (id | 최대 글자 | 줄 수 | 지금 글)\n" + "\n".join(rows)
         + '\n\n# 출력\nJSON 하나만: {"texts": {"id": "줄인 글", ...}}')
    open(outp, "w", encoding="utf-8").write(p)
    print("줄일 항목", len(rows), "도형", len(want))


APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
       'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office PowerPoint</Application><Company>Human Consulting Group</Company></Properties>')
CORE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>@@TITLE@@</dc:title><dc:creator>HCG</dc:creator><cp:lastModifiedBy>HCG</cp:lastModifiedBy>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-06T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-11-06T00:00:00Z</dcterms:modified></cp:coreProperties>')


def props(src, dst):
    """파일 속성(장 제목 목록 · 작성자 · 옛 표지 미리보기)을 새로 쓰고, 다른 고객 이름이 남았는지 꾸러미 전체를 훑음."""
    zi = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            b = zi.read(it.filename)
            if it.filename == "docProps/app.xml": b = APP.encode()
            if it.filename == "docProps/core.xml": b = CORE.replace("@@TITLE@@", f"{CLIENT} {TITLE} 제안서 Draft").encode()
            if it.filename == "docProps/thumbnail.jpeg": continue
            if it.filename == "_rels/.rels":
                b = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b"", b)
            if it.filename == "[Content_Types].xml":
                b = re.sub(rb'<Override [^>]*thumbnail[^>]*/>', b"", b)
            zo.writestr(it, b)
    names = other_names()
    z = zipfile.ZipFile(dst); hits = {}
    skip_parts = keep_parts(z)
    for nm in z.namelist():
        if nm in skip_parts:
            continue
        if nm.endswith((".xml", ".rels")):
            txt = " ".join(re.findall(r">([^<>]+)<", z.read(nm).decode("utf-8", "ignore")))
            h = [x for x in names if x in txt]
            if h: hits[nm] = h
        elif nm.startswith("ppt/embeddings/") and nm.endswith(".xlsx"):
            import io
            zz = zipfile.ZipFile(io.BytesIO(z.read(nm)))
            txt = " ".join(zz.read(x).decode("utf-8", "ignore") for x in zz.namelist() if x.endswith(".xml"))
            h = [x for x in names if x in txt]
            if h: hits[nm] = h
    print("다른 고객 이름", hits or 0)
    return hits


def keep_parts(z):
    """그대로 넣은 장의 장 XML 과 거기서 닿는 차트 · 엑셀(HCG 소개의 고객 이름은 정상)."""
    skip = keep_slides()
    if not skip:
        return set()
    import posixpath
    def rels(part):
        d, f = posixpath.split(part)
        rn = posixpath.join(d, "_rels", f + ".rels")
        if rn not in z.namelist():
            return []
        x = z.read(rn).decode("utf-8", "ignore")
        return [posixpath.normpath(posixpath.join(d, t)) for t, m in re.findall(r'Target="([^"]+)"(?:\s+TargetMode="(\w+)")?', x) if m != "External"]
    pres = z.read("ppt/presentation.xml").decode("utf-8", "ignore")
    prel = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "ignore")))
    prel.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "ignore"))})
    order = [posixpath.normpath(posixpath.join("ppt", prel[r])) for r in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', pres) if r in prel]
    out = set()
    for i in skip:
        if i - 1 < len(order):
            sp = order[i - 1]
            out.add(sp)
            out |= {t for t in rels(sp) if "/charts/" in t or "/embeddings/" in t}
    return out


def other_names():
    """Library 의 고객 이름(옛 이름 포함) + 틀 제안서 고객 — 이번 고객 이름은 뺌."""
    meta = json.load(open(META, encoding="utf-8"))
    names = set()
    for r in meta:
        for f in (r.get("olds") or []) + [r.get("client") or ""]:
            f = re.sub(r"(㈜|\(주\)|주식회사)", "", f).strip()
            if len(f) >= 3 and f != "HCG": names.add(f)
    names |= {"쿠리타", "KURITA", "Kurita", "주성", "제테마", "한국카본", "NICE", "코오롱"}
    return {x for x in names if not any(c in x or x in c for c in CLIENT_NAMES)}


def iter_paras(prs, skip=()):
    """장 안 모든 글 문단(표 칸 포함)을 (항목 id, 문단) 으로."""
    for i, s in enumerate(prs.slides, 1):
        if i in skip:
            continue
        for sh, _ in walk_abs(s.shapes):
            if sh.has_text_frame:
                for k, para in enumerate(sh.text_frame.paragraphs):
                    yield f"{i}:{sh.shape_id}:{k}", para
            elif getattr(sh, "has_table", False) and sh.has_table:
                for r, row in enumerate(sh.table.rows):
                    for c, cell in enumerate(row.cells):
                        for k, para in enumerate(cell.text_frame.paragraphs):
                            yield f"{i}:{sh.shape_id}:r{r}c{c}:{k}", para


def scan_names(deck):
    """장 글에 남은 다른 고객 이름 → [(항목 id, 글, [이름])]."""
    names = other_names()
    out = []
    for key, para in iter_paras(Presentation(deck), keep_slides()):
        t = ptext(para)
        h = [x for x in names if x in t]
        if h:
            out.append((key, t, h))
    return out


if __name__ == "__main__":
    globals()[sys.argv[1]](*sys.argv[2:])
