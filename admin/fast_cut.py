# 빠른 자르기 — 원본 PPTX 를 문서당 한 번만 읽고, 필요한 장마다 「그 장 + 쓰는 레이아웃 · 마스터 · 테마」만 담은 한 장짜리 PPTX 를 zip 수준에서 만듦
# (store_build.cut_slide 는 장마다 원본을 통째로 다시 열어 큰 파일에서 느림)
# 실행: ..\..\.venv\Scripts\python -X utf8 fast_cut.py --grades 상,중     → 아직 없는 장만 store/out/pptx 에
import argparse, collections, json, posixpath, re, time, zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree
from pptx import Presentation

import scrub
import store_build as B

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
RID = "{%s}id" % NS["r"]
T_SLIDE = "/slide"
DROP_FROM_PRES = ("/slide", "/slideMaster", "/font", "/customShow")  # 남길 것만 다시 넣음


def rels_name(part):
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def read_rels(z, part):
    name = rels_name(part)
    if name not in z:
        return []
    root = etree.fromstring(z[name])
    out = []
    for r in root:
        tgt = r.get("Target")
        mode = r.get("TargetMode")
        full = tgt if mode == "External" else posixpath.normpath(posixpath.join(posixpath.dirname(part), tgt)).lstrip("/")
        out.append({"id": r.get("Id"), "type": r.get("Type"), "target": tgt, "full": full, "ext": mode == "External"})
    return out


def write_rels(rels):
    root = etree.Element("{%s}Relationships" % NS["rel"], nsmap={None: NS["rel"]})
    for r in rels:
        e = etree.SubElement(root, "{%s}Relationship" % NS["rel"])
        e.set("Id", r["id"]); e.set("Type", r["type"]); e.set("Target", r["target"])
        if r["ext"]:
            e.set("TargetMode", "External")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class Deck:
    def __init__(self, path):
        with zipfile.ZipFile(path) as zf:
            self.z = {n: zf.read(n) for n in zf.namelist()}
        self.pres = "ppt/presentation.xml"
        self.pres_rels = read_rels(self.z, self.pres)
        root = etree.fromstring(self.z[self.pres])
        byid = {r["id"]: r for r in self.pres_rels}
        self.slides = [byid[s.get(RID)]["full"] for s in root.find("p:sldIdLst", NS)]

    def cut(self, i, out_path):
        z = self.z
        slide = self.slides[i - 1]
        srels = read_rels(z, slide)
        layout = next(r["full"] for r in srels if r["type"].endswith("/slideLayout"))
        lrels = read_rels(z, layout)
        master = next(r["full"] for r in lrels if r["type"].endswith("/slideMaster"))
        files = {}
        # presentation.xml — 장 하나 · 마스터 하나만, 심은 글꼴 · 사용자 지정 쇼 · 구역 빼기
        proot = etree.fromstring(z[self.pres])
        byfull = {}
        keep_rels = []
        for r in self.pres_rels:
            t = r["type"]
            if t.endswith("/slide"):
                if r["full"] == slide:
                    keep_rels.append(r); byfull["slide"] = r["id"]
            elif t.endswith("/slideMaster"):
                if r["full"] == master:
                    keep_rels.append(r); byfull["master"] = r["id"]
            elif t.endswith("/font"):
                continue
            else:
                keep_rels.append(r)
        sl = proot.find("p:sldIdLst", NS)
        for s in list(sl):
            if s.get(RID) != byfull["slide"]:
                sl.remove(s)
        ml = proot.find("p:sldMasterIdLst", NS)
        for m in list(ml):
            if m.get(RID) != byfull["master"]:
                ml.remove(m)
        for tag in ("p:embeddedFontLst", "p:custShowLst"):
            for e in proot.findall(tag, NS):
                proot.remove(e)
        for e in proot.iter("{%s}extLst" % NS["p"]):
            for ext in list(e):
                if any(c.tag.endswith("}sectionLst") for c in ext):
                    e.remove(ext)
        files[self.pres] = etree.tostring(proot, xml_declaration=True, encoding="UTF-8", standalone=True)
        files[rels_name(self.pres)] = write_rels(keep_rels)
        # 마스터 — 쓰는 레이아웃만
        mroot = etree.fromstring(z[master])
        mrels = read_rels(z, master)
        mby = {r["id"]: r for r in mrels}
        ll = mroot.find("p:sldLayoutIdLst", NS)
        drop = set()
        if ll is not None:
            for lid in list(ll):
                r = mby.get(lid.get(RID))
                if not r or r["full"] != layout:
                    ll.remove(lid)
                    if r:
                        drop.add(r["id"])
        files[master] = etree.tostring(mroot, xml_declaration=True, encoding="UTF-8", standalone=True)
        files[rels_name(master)] = write_rels([r for r in mrels if r["id"] not in drop])
        # 장 — 다른 장으로 가는 링크 끊기
        sroot = etree.fromstring(z[slide])
        bad = {r["id"] for r in srels if r["type"].endswith("/slide")}
        if bad:
            for el in list(sroot.iter()):
                if any(k == RID and v in bad for k, v in el.attrib.items()) and el.getparent() is not None:
                    el.getparent().remove(el)
            files[slide] = etree.tostring(sroot, xml_declaration=True, encoding="UTF-8", standalone=True)
            files[rels_name(slide)] = write_rels([r for r in srels if r["id"] not in bad])
        # 닿는 부분만 모음(관계를 따라감, 다른 장 · 장 메모의 다른 장은 안 감)
        def rels_of(part):
            name = rels_name(part)
            if name in files:
                root = etree.fromstring(files[name])
                return [{"type": r.get("Type"), "full": r.get("Target") if r.get("TargetMode") == "External" else posixpath.normpath(posixpath.join(posixpath.dirname(part), r.get("Target"))).lstrip("/"), "ext": r.get("TargetMode") == "External"} for r in root]
            return read_rels(z, part)
        keep = set()
        stack = [r["full"] for r in read_rels(z, "") if not r["ext"]] if "_rels/.rels" in z else []
        stack.append(self.pres)
        while stack:
            part = stack.pop()
            if part in keep or part not in z:
                continue
            keep.add(part)
            for r in rels_of(part):
                if r["ext"]:
                    continue
                if r["type"].endswith("/slide") and r["full"] != slide:
                    continue  # 메모 장 → 다른 장 등
                stack.append(r["full"])
        keep |= {"[Content_Types].xml", "_rels/.rels"}
        names = set()
        for p in keep:
            names.add(p)
            rn = rels_name(p)
            if rn in z or rn in files:
                names.add(rn)
        # 내용 종류 — 없는 부분 지우기
        ct = etree.fromstring(z["[Content_Types].xml"])
        for o in list(ct):
            if o.tag.endswith("}Override") and o.get("PartName").lstrip("/") not in keep:
                ct.remove(o)
        files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in ["[Content_Types].xml"] + sorted(names - {"[Content_Types].xml"}):
                data = files.get(n, z.get(n))
                if data is not None:
                    zout.writestr(n, data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grades", default="상,중")
    a = ap.parse_args()
    grades = set(a.grades.split(",")) | ({"中"} if "중" in a.grades else set())
    idx, conv, refs = B.load_index()
    names = B.known_names()
    price = re.compile(r"가격제안|견적|04 가격")
    rows = [json.loads(l) for l in open(B.WORK / "ref" / "review" / "tags_all.jsonl", encoding="utf-8")]
    need = [r for r in rows if r["rep"] == r["k"] and (r["re"] in grades or r["star"]) and r["p"] != "P26"
            and not price.search(r["name"]) and not (B.OUT / "pptx" / f"s{r['k']:05d}.pptx").exists()]
    by = collections.defaultdict(list)
    for r in need:
        by[r["doc"]].append(r)
    print("자를 장", len(need), "· 문서", len(by), flush=True)
    t0, done, fail = time.time(), 0, []
    for d, lst in sorted(by.items(), key=lambda kv: -len(kv[1])):
        src = B.source_path(lst[0], idx, conv, refs)
        if not src:
            fail += [(f"s{r['k']:05d}", "원본 없음") for r in lst]
            continue
        try:
            deck = Deck(src)
        except Exception as e:
            fail += [(f"s{r['k']:05d}", f"원본 열기 {e!r}") for r in lst]
            continue
        for r in lst:
            sid = f"s{r['k']:05d}"
            pp = B.OUT / "pptx" / f"{sid}.pptx"
            try:
                deck.cut(r["i"], pp)
                scrub.scrub_file(pp, names)
                assert len(Presentation(str(pp)).slides) == 1
                done += 1
            except Exception as e:
                if pp.exists():
                    pp.unlink()
                try:  # 느린 길로 한 번 더
                    B.cut_slide(src, r["i"], pp)
                    scrub.scrub_file(pp, names)
                    done += 1
                except Exception as e2:
                    fail.append((sid, f"{e!r} / {e2!r}"))
        print(f"{len(lst):4d}장 {time.time() - t0:6.0f}s · 누적 {done} · {Path(str(src)).name[:50]}", flush=True)
    print("끝", done, "장 · 실패", len(fail), f"· {time.time() - t0:.0f}s")
    for x in fail[:20]:
        print("  실패", *x)


if __name__ == "__main__":
    main()
