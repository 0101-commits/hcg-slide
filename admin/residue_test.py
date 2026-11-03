# 2단계 확인 — 「넣은 장표에 이전 고객명이 남지 않음」
# 작업창(pane.js)과 같은 방식으로 원본 고객명을 새 고객명으로 바꾼 뒤, PPTX 안 모든 XML 에서 원본 고객명이 남았는지 찾음.
# 실행: ..\..\.venv\Scripts\python -X utf8 residue_test.py [--new 한빛소재]
import argparse, collections, json, re, zipfile
from pathlib import Path

from lxml import etree

OUT = Path(__file__).resolve().parent.parent / "store" / "out"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
RX = re.compile(r"^ppt/(slides/slide\d+|notesSlides/notesSlide\d+|slideMasters/slideMaster\d+|slideLayouts/slideLayout\d+|charts/chart\d+|diagrams/(data|drawing)\d+)\.xml$")
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def replace_xml(data, pairs):
    root = etree.fromstring(data)
    n = 0
    for p in root.iter(f"{{{A}}}p"):
        ts = list(p.iter(f"{{{A}}}t"))
        if not ts:
            continue
        for frm, to in pairs:
            full = "".join(t.text or "" for t in ts)
            at = full.find(frm)
            while at >= 0:
                pos, first, end = 0, True, at + len(frm)
                for t in ts:
                    s = t.text or ""
                    a, b = pos, pos + len(s)
                    pos = b
                    lo, hi = max(a, at), min(b, end)
                    if lo >= hi:
                        continue
                    t.text = s[: lo - a] + (to if first else "") + s[hi - a:]
                    first = False
                n += 1
                full = "".join(t.text or "" for t in ts)
                at = full.find(frm, at + len(to))
    for v in root.iter(f"{{{C}}}v"):  # 차트에 담긴 글 값
        if v.text:
            nv = v.text
            for frm, to in pairs:
                nv = nv.replace(frm, to)
            if nv != v.text:
                v.text = nv
                n += 1
    for el in root.iter():
        if isinstance(el.tag, str) and el.tag.endswith("}cNvPr"):
            for att in ("name", "descr", "title"):
                v = el.get(att)
                if v:
                    nv = v
                    for frm, to in pairs:
                        nv = nv.replace(frm, to)
                    if nv != v:
                        el.set(att, nv)
                        n += 1
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--new", default="한빛소재")
    a = ap.parse_args()
    meta = json.load(open(OUT / "meta.json", encoding="utf-8"))
    tested = clean = 0
    left = collections.Counter()
    bad = []
    for m in meta:
        c = m.get("client")
        if not c or c not in (m.get("olds") or []):
            continue  # 원본 고객명이 장표 글에 없는 장(작업창이 미리 채우지 않음)
        if m.get("p") in ("P22", "P23"):
            continue  # 실적 · 투입 인력 장표는 미리 채우지 않음(실적 목록일 수 있음)
        tested += 1
        z = zipfile.ZipFile(OUT / "pptx" / f"{m['id']}.pptx")
        where = []
        for name in z.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            data = z.read(name)
            if RX.match(name):
                data, _ = replace_xml(data, [(c, a.new)])
            if c.encode("utf-8") in data:
                where.append(name)
        if where:
            kinds = sorted({re.sub(r"\d+", "N", w) for w in where})
            bad.append((m["id"], c, kinds))
            for k in kinds:
                left[k] += 1
        else:
            clean += 1
    print(f"원본 고객명이 글에 있는 장 {tested} · 바꾼 뒤 남지 않음 {clean} · 남음 {len(bad)}")
    for k, v in left.most_common():
        print(f"  남은 곳 {k}: {v}장")
    for x in bad[:15]:
        print("  ", *x)


if __name__ == "__main__":
    main()
