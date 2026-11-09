# 한 장 PPTX 안의 「닿지 않는 부품」 정리 — 잘라낸 다른 장이 viewProps(마지막으로 본 장) 연결로 숨어 남으면
# PowerPoint 가 파일을 못 열고(손상), 다른 장 글 · 그림이 파일 안에 남음. _rels/.rels 에서 관계를 따라 닿는 부품만 남김.
import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree


def _rels_of(part):
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def reachable(out):
    seen, todo = set(), [""]
    while todo:
        part = todo.pop()
        rn = "_rels/.rels" if part == "" else _rels_of(part)
        if rn not in out:
            continue
        base = posixpath.dirname(part)
        for rel in etree.fromstring(out[rn]):
            if rel.get("TargetMode") == "External":
                continue
            t = rel.get("Target", "")
            p = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
            if p in out and p not in seen:
                seen.add(p)
                todo.append(p)
    return seen


def clean(out):
    """out = {이름: 바이트}. viewProps 빼고 닿지 않는 부품 · 관계 · 형식 등록을 지움. 지운 이름 목록."""
    removed = []
    if "ppt/viewProps.xml" in out:  # 마지막으로 본 장 · 개요 보기 장 목록 — PowerPoint 가 다시 만듦
        for n in ("ppt/viewProps.xml", "ppt/_rels/viewProps.xml.rels"):
            if n in out:
                del out[n]
                removed.append(n)
        pr = "ppt/_rels/presentation.xml.rels"
        out[pr] = re.sub(rb'<Relationship [^>]*Target="viewProps\.xml"[^>]*/>', b"", out[pr])
    keep = reachable(out)
    for n in list(out):
        if n == "[Content_Types].xml" or n == "_rels/.rels":
            continue
        if n.endswith(".rels"):
            src = n.replace("_rels/", "")[:-5]
            if src in keep:
                continue
        elif n in keep:
            continue
        del out[n]
        removed.append(n)
    ct = etree.fromstring(out["[Content_Types].xml"])
    for o in list(ct):
        if o.tag.endswith("Override") and o.get("PartName", "").lstrip("/") not in out:
            ct.remove(o)
    out["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    return removed


def clean_file(path, dry=False):
    path = Path(path)
    zi = zipfile.ZipFile(path)
    out = {n: zi.read(n) for n in zi.namelist()}
    zi.close()
    # viewProps 가 지금 장 목록에 없는(잘려 나간) 장을 가리키면 PowerPoint 가 못 엶
    vp_bad = False
    vr = "ppt/_rels/viewProps.xml.rels"
    if vr in out:
        pres = out.get("ppt/_rels/presentation.xml.rels", b"").decode("utf-8", "ignore")
        listed = {posixpath.normpath(posixpath.join("ppt", t)) for t in re.findall(r'Target="(slides/slide\d+\.xml)"', pres)}
        for t in re.findall(r'Target="([^"]*slide\d+\.xml)"', out[vr].decode("utf-8", "ignore")):
            if posixpath.normpath(posixpath.join("ppt", t)) not in listed:
                vp_bad = True
    removed = clean(out)
    # viewProps 만 빠지는 멀쩡한 파일은 그대로 둠(다시 올리지 않게) — 숨은 부품 · 잘못된 viewProps 가 있을 때만 다시 씀
    real = [n for n in removed if not n.startswith(("ppt/viewProps", "ppt/_rels/viewProps"))] or (removed if vp_bad else [])
    if real and not dry:
        tmp = path.with_suffix(".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", out.pop("[Content_Types].xml"))
            for n, d in out.items():
                z.writestr(n, d)
        tmp.replace(path)
    return removed if real else []


if __name__ == "__main__":
    import sys
    from collections import Counter
    dry = "--dry" in sys.argv
    files = sorted((Path(__file__).resolve().parent.parent / "store" / "out" / "pptx").glob("*.pptx"))
    hit, kinds = 0, Counter()
    for f in files:
        r = clean_file(f, dry=dry)
        if r:
            hit += 1
            for n in r:
                kinds[re.sub(r"\d+", "#", n)] += 1
    print(("(시험) " if dry else "") + f"파일 {len(files)} · 고친 파일(숨은 부품 · 잘못된 viewProps) {hit}")
    for k, v in kinds.most_common(12):
        print(" ", v, k)
