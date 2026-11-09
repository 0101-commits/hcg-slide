# 한 장 PPTX 다듬기 — 이전 고객 정보가 숨은 곳을 지움(store_build 가 자를 때 부르고, 단독으로 이미 만든 파일에도 돌림)
#  · docProps/app.xml : 원본 파일의 장 제목 목록(TitlesOfParts) · 회사 · 관리자 → 빈 기본값
#  · 테마 이름 · 마스터 · 레이아웃 이름에 고객명이 있으면 「HCG」로
# 실행: ..\..\.venv\Scripts\python -X utf8 scrub.py   (store/out/pptx 전체)
import os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

APP_XML = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           b'<Application>Microsoft Office PowerPoint</Application><Slides>1</Slides><Company>HCG</Company></Properties>')
NAME_RX = re.compile(rb'(<a:theme\b[^>]*\bname=")([^"]*)(")|(<p:cSld\b[^>]*\bname=")([^"]*)(")')


CORE_TAGS = re.compile(rb"(<(dc:title|dc:subject|cp:keywords|dc:description|cp:category|dc:creator|cp:lastModifiedBy|cp:contentStatus)\b[^>]*>)(.*?)(</\2>)", re.S)
THEME_NAME = re.compile(rb'(\bname=")([^"]*)(")')


def scrub_bytes(name, data, names):
    if name == "docProps/app.xml":
        return APP_XML
    if name == "docProps/core.xml":  # 제목 · 주제 · 작성자 등 — 원본 제안서 정보
        return CORE_TAGS.sub(lambda m: m.group(1) + (b"HCG" if m.group(2) in (b"dc:creator", b"cp:lastModifiedBy") else b"") + m.group(4), data)
    if name.startswith("ppt/theme/") and name.endswith(".xml"):  # 테마 · 색 · 글꼴 묶음 이름
        def fixn(m):
            v = m.group(2).decode("utf-8", "ignore")
            return m.group(1) + (b"HCG" if any(n in v for n in names) else m.group(2)) + m.group(3)
        data = THEME_NAME.sub(fixn, data)
    if name.startswith(("ppt/theme/", "ppt/slideMasters/", "ppt/slideLayouts/")) and name.endswith(".xml"):
        def fix(m):
            pre, val, post = (m.group(1), m.group(2), m.group(3)) if m.group(1) else (m.group(4), m.group(5), m.group(6))
            v = val.decode("utf-8", "ignore")
            if any(n in v for n in names):
                v = "HCG"
            return pre + v.encode("utf-8") + post
        return NAME_RX.sub(fix, data)
    return data


CORE_XML = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            b'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            b'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            b'<dc:creator>HCG</dc:creator><cp:lastModifiedBy>HCG</cp:lastModifiedBy></cp:coreProperties>')
# 원본 덱 첫 장 미리보기 그림 · 사용자 속성 — 다른 고객 표지가 보일 수 있어 아예 뺌
DROP = re.compile(r"^docProps/(thumbnail\.[a-z]+|custom\.xml)$", re.I)


def scrub_file(path, names):
    path = Path(path)
    zin = zipfile.ZipFile(path)
    fd, tmpn = tempfile.mkstemp(suffix=".tmp", dir=path.parent)
    os.close(fd)
    tmp = Path(tmpn)
    changed = False
    dropped = {n for n in zin.namelist() if DROP.match(n)}
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            if it.filename in dropped:
                changed = True
                continue
            data = zin.read(it.filename)
            if it.filename == "docProps/core.xml":
                new = CORE_XML  # 접두어와 상관없이 통째로 새로
            elif it.filename == "_rels/.rels" and dropped:
                new = re.sub(rb'<Relationship\b[^>]*Target="/?docProps/(thumbnail\.[a-zA-Z]+|custom\.xml)"[^>]*/>', b"", data)
            elif it.filename == "[Content_Types].xml" and dropped:
                new = re.sub(rb'<Override\b[^>]*PartName="/docProps/(thumbnail\.[a-zA-Z]+|custom\.xml)"[^>]*/>', b"", data)
            else:
                new = scrub_bytes(it.filename, data, names)
            changed |= new != data
            zout.writestr(it, new)
    zin.close()
    if changed:
        shutil.move(str(tmp), str(path))
    else:
        tmp.unlink()
    # 잘려 나간 장이 viewProps 연결로 숨어 남으면 PowerPoint 가 못 엶 + 다른 장 내용이 남음 → 닿지 않는 부품 정리(11-04)
    import orphans
    changed |= bool(orphans.clean_file(path))
    return changed


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    import store_build as B
    names = B.known_names()
    out = B.OUT / "pptx"
    n = sum(scrub_file(p, names) for p in sorted(out.glob("*.pptx")))
    print("다듬음", n, "/", len(list(out.glob("*.pptx"))))
