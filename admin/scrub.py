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


def scrub_bytes(name, data, names):
    if name == "docProps/app.xml":
        return APP_XML
    if name.startswith(("ppt/theme/", "ppt/slideMasters/", "ppt/slideLayouts/")) and name.endswith(".xml"):
        def fix(m):
            pre, val, post = (m.group(1), m.group(2), m.group(3)) if m.group(1) else (m.group(4), m.group(5), m.group(6))
            v = val.decode("utf-8", "ignore")
            if any(n in v for n in names):
                v = "HCG"
            return pre + v.encode("utf-8") + post
        return NAME_RX.sub(fix, data)
    return data


def scrub_file(path, names):
    path = Path(path)
    zin = zipfile.ZipFile(path)
    fd, tmpn = tempfile.mkstemp(suffix=".tmp", dir=path.parent)
    os.close(fd)
    tmp = Path(tmpn)
    changed = False
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            new = scrub_bytes(it.filename, data, names)
            changed |= new != data
            zout.writestr(it, new)
    zin.close()
    if changed:
        shutil.move(str(tmp), str(path))
    else:
        tmp.unlink()
    return changed


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    import store_build as B
    names = B.known_names()
    out = B.OUT / "pptx"
    n = sum(scrub_file(p, names) for p in sorted(out.glob("*.pptx")))
    print("다듬음", n, "/", len(list(out.glob("*.pptx"))))
