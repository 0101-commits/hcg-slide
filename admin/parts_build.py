# 3단계 부품 12개 — HCG 템플릿 파일(표준 템플릿 · Time Saver · Graphic · Ghost Pack)에서 구조별 후보 3개를 골라 한 장짜리로 자름
# 실행: ..\..\.venv\Scripts\python -X utf8 parts_build.py   → store/out/{pptx,img,thumb}/s#####.* + pane/public/pane/parts.json + parts_sheet.jpg(확인용)
import json, re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import scrub
import store_build as B

PARTS = [("T1", "표(일반)"), ("L2", "3~4단 기둥"), ("F1", "가로 단계"), ("L1", "좌우 2단"), ("L5", "좌우 비교(전→후)"), ("C2", "옵션 비교표"),
         ("F5", "간트"), ("H1", "조직도"), ("H2", "로직 트리"), ("H6", "중심-방사"), ("H5", "프레임워크 블록"), ("D1", "차트 + 시사점")]
DOC_PREF = {"80c809c6e3": 0, "cc8b7985fa": 1, "41820cb8c8": 2, "1ead053e91": 3}  # 표준 템플릿 > Time Saver > Graphic > Ghost Pack
GRADE = {"상": 0, "중": 1, "中": 1, "하": 2}
PER = 3


def main():
    idx, conv, refs = B.load_index()
    names = B.known_names()
    rows = [json.loads(l) for l in open(B.WORK / "ref" / "review" / "tags_all.jsonl", encoding="utf-8")]
    slides = {}
    for l in open(B.WORK / "ref" / "review" / "slides_all.jsonl", encoding="utf-8"):
        d = json.loads(l)
        slides[d["k"]] = d
    uniq = [r for r in rows if r["rep"] == r["k"] and r["doc"] in DOC_PREF]
    out, sheet = [], []
    for code, name in PARTS:
        # 템플릿 사용 안내(아이콘 · 그림 형식 설명)와 글 없는 도형 모음은 부품이 아님
        guide = re.compile(r"SVG|PNG|JPEG|아이콘의|색상 적용법|사용법|활용법|가이드|Guide")
        cand = [r for r in uniq if r["s"] == code
                and not guide.search((r.get("title") or "") + " " + (slides.get(r["k"], {}).get("text") or ""))
                and len(slides.get(r["k"], {}).get("text") or "") >= 30]
        # 글이 아주 많거나(안내문) 거의 없는 장은 뒤로
        def score(r):
            n = len((slides.get(r["k"], {}).get("text") or ""))
            return (GRADE.get(r["re"], 3), DOC_PREF[r["doc"]], 0 if 40 <= n <= 900 else 1, r["k"])
        cand.sort(key=score)
        vars_ = []
        for r in cand:
            if len(vars_) >= PER:
                break
            src = B.source_path(r, idx, conv, refs)
            if not src:
                continue
            sid = f"s{r['k']:05d}"
            pp = B.OUT / "pptx" / f"{sid}.pptx"
            if not pp.exists():
                B.cut_slide(src, r["i"], pp)
                scrub.scrub_file(pp, names)
            im = Image.open(r["png"]).convert("RGB")
            for d, w, q in (("img", 960, 84), ("thumb", 360, 80)):
                f = B.OUT / d / f"{sid}.jpg"
                if not f.exists():
                    im.resize((w, round(im.height * w / im.width)), Image.LANCZOS).save(f, quality=q, optimize=True)
            title = " ".join((r.get("title") or "").split())[:60]
            vars_.append({"id": sid, "title": title, "doc": re.split(r"[\\/]", r["name"])[-1], "grade": r["re"]})
            sheet.append((code, name, sid, B.OUT / "thumb" / f"{sid}.jpg"))
        out.append({"code": code, "name": name, "struct": B.STRUCT.get(code, code), "count": len([r for r in uniq if r["s"] == code]), "variants": vars_})
        print(code, name, [v["id"] for v in vars_])
    dst = B.APP / "pane" / "public" / "pane" / "parts.json"
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    # 확인용 그림(사람이 한 번 봄)
    W, H = 360, 203
    cols = PER
    img = Image.new("RGB", (cols * (W + 8) + 8, len(PARTS) * (H + 26) + 8), "white")
    dr = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 14)
    except OSError:
        font = ImageFont.load_default()
    rowi = {c: i for i, (c, _) in enumerate(PARTS)}
    colc = {}
    for code, name, sid, f in sheet:
        i = rowi[code]
        j = colc.get(code, 0)
        colc[code] = j + 1
        x, y = 8 + j * (W + 8), 8 + i * (H + 26)
        t = Image.open(f)
        t.thumbnail((W, H))
        img.paste(t, (x, y + 20))
        dr.text((x, y + 2), f"{code} {name} · {sid}", fill="black", font=font)
    img.save(B.APP / "store" / "parts_sheet.jpg", quality=80)
    print("parts.json", len(out), "부품 ·", sum(len(p["variants"]) for p in out), "모양")


if __name__ == "__main__":
    main()
