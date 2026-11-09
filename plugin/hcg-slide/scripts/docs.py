# 고객 자료 폴더 읽기 — 파일을 읽기만 하고(원본 폴더에 아무것도 쓰지 않음) 글을 뽑아 작업 폴더의 docs.md 로 모은다.
# 읽는 것: docx · doc(Word 필요) · pptx · xlsx · pdf(스캔본은 Windows OCR) · hwp · hwpx · txt · md · csv. 그림 파일 · 동영상 · 압축 파일은 건너뜀.
import os, re, struct, zipfile, zlib

def stype(sh):
    """도형 종류. python-pptx 가 모르는 도형(특수 자동 도형)은 None."""
    try:
        return sh.shape_type
    except NotImplementedError:
        return None

PER_FILE = 9000      # 파일 하나에서 가져올 글자 수
TOTAL = 60000        # 모두 합쳐 가져올 글자 수
UNZIP_CAP = 60 * 1024 * 1024    # 압축 안 글 하나가 풀렸을 때 상한(압축 폭탄 막기)
UNZIP_TOTAL = 400 * 1024 * 1024  # 파일 하나 전체가 풀렸을 때 상한
EXTS = (".docx", ".doc", ".pptx", ".xlsx", ".pdf", ".hwp", ".hwpx", ".txt", ".md", ".csv")


def _zip(path):
    z = zipfile.ZipFile(path)
    if sum(i.file_size for i in z.infolist()) > UNZIP_TOTAL:
        raise RuntimeError("풀면 너무 큰 파일 — 건너뜀")
    return z


def _zread(z, name):
    if z.getinfo(name).file_size > UNZIP_CAP:
        raise RuntimeError(f"풀면 너무 큰 부분({name}) — 건너뜀")
    return z.read(name)


def _xml_text(x, para_tag, text_tag):
    out = []
    for para in re.findall(rf"<{para_tag}[ >].*?</{para_tag}>", x, re.S):
        ts = re.findall(rf"<{text_tag}[^>]*>(.*?)</{text_tag}>", para, re.S)
        t = "".join(re.sub(r"<[^>]+>", "", s) for s in ts).strip()
        if t:
            out.append(t)
    return out


def _unesc(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def read_docx(path):
    z = _zip(path)
    x = _zread(z, "word/document.xml").decode("utf-8", "replace")
    lines = []
    # 표는 줄마다 칸을 | 로 이음
    pos = 0
    for m in re.finditer(r"<w:tbl>.*?</w:tbl>", x, re.S):
        lines += _xml_text(x[pos:m.start()], "w:p", "w:t")
        lines.append("[표]")
        for row in re.findall(r"<w:tr[ >].*?</w:tr>", m.group(0), re.S):
            cells = [" ".join(_xml_text(c, "w:p", "w:t")) for c in re.findall(r"<w:tc>.*?</w:tc>", row, re.S)]
            lines.append(" | ".join(cells))
        pos = m.end()
    lines += _xml_text(x[pos:], "w:p", "w:t")
    return "\n".join(_unesc(l) for l in lines)


def read_pptx(path):
    from pptx import Presentation
    _zip(path).close()
    prs = Presentation(path)
    out = []

    def walk(shs):
        for sh in shs:
            if stype(sh) == 6:
                yield from walk(sh.shapes)
            else:
                yield sh
    for i, s in enumerate(prs.slides, 1):
        out.append(f"[장 {i}]")
        for sh in walk(s.shapes):
            if sh.has_text_frame and sh.text_frame.text.strip():
                out.append(sh.text_frame.text.strip().replace("\n", " / "))
            if getattr(sh, "has_table", False) and sh.has_table:
                for r in sh.table.rows:
                    out.append(" | ".join(c.text.strip().replace("\n", " ") for c in r.cells))
    return "\n".join(out)


def read_xlsx(path):
    z = _zip(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        x = _zread(z, "xl/sharedStrings.xml").decode("utf-8", "replace")
        for si in re.findall(r"<si>.*?</si>", x, re.S):
            shared.append(_unesc("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    names = re.findall(r'<sheet [^>]*name="([^"]+)"', _zread(z, "xl/workbook.xml").decode("utf-8", "replace"))
    out = []
    sheets = sorted((n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n)), key=lambda n: int(re.findall(r"\d+", n)[-1]))
    for k, n in enumerate(sheets):
        out.append(f"[시트 {names[k] if k < len(names) else k + 1}]")
        x = _zread(z, n).decode("utf-8", "replace")
        for r_i, row in enumerate(re.findall(r"<row[ >].*?</row>", x, re.S)):
            if r_i >= 300:
                break
            vals = []
            for c in re.findall(r"<c [^>]*?(?:/>|>.*?</c>)", row, re.S):
                v = re.search(r"<v>(.*?)</v>", c, re.S)
                t = re.search(r"<t[^>]*>(.*?)</t>", c, re.S)
                if 't="s"' in c and v:
                    try:
                        vals.append(shared[int(v.group(1))])
                    except (ValueError, IndexError):
                        pass
                elif t:
                    vals.append(_unesc(t.group(1)))
                elif v:
                    vals.append(v.group(1))
            vals = [s.strip() for s in vals if s and s.strip()]
            if vals:
                out.append(" | ".join(vals[:40]))
    return "\n".join(out)


def read_pdf(path):
    from pypdf import PdfReader
    r = PdfReader(path)
    out = []
    for i, p in enumerate(r.pages, 1):
        if i > 60:
            break
        try:
            t = (p.extract_text() or "").strip()
        except Exception:
            t = ""
        if t:
            out.append(f"[쪽 {i}]\n{t}")
    if not out:  # 스캔본 → Windows 내장 OCR(앞 15쪽)
        t = _ps("pdfocr.ps1", path)
        if not t.strip():
            raise RuntimeError("글이 없는 PDF(스캔본) — OCR 로도 못 읽음")
        return t
    return "\n".join(out)


def _ps(script, path):
    """같은 폴더의 PowerShell 도우미로 파일 글을 뽑음(임시 글 파일 경유). 실패하면 빈 글."""
    import subprocess, tempfile
    fd, tmp = tempfile.mkstemp(suffix=".txt"); os.close(fd)
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", os.path.join(os.path.dirname(os.path.abspath(__file__)), script),
                        "-File", path, "-Out", tmp], capture_output=True, timeout=600, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return open(tmp, encoding="utf-8").read() if os.path.getsize(tmp) else ""
    except Exception:
        return ""
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def read_doc(path):
    t = _ps("doc2txt.ps1", path)
    if not t.strip():
        raise RuntimeError("옛 Word(.doc) — 이 PC 에 Word 가 없거나 못 열음")
    return t


def read_hwp(path):
    import olefile
    ole = olefile.OleFileIO(path)
    try:
        hdr = ole.openstream("FileHeader").read()
        compressed = bool(hdr[36] & 1)
        secs = sorted((e for e in ole.listdir() if e[0] == "BodyText"), key=lambda e: int(e[1].replace("Section", "")))
        out = []
        for e in secs:
            data = ole.openstream(e).read()
            if compressed:
                data = zlib.decompressobj(-15).decompress(data, UNZIP_CAP)
            i = 0
            while i + 4 <= len(data):
                h = struct.unpack_from("<I", data, i)[0]
                tag, size = h & 0x3FF, (h >> 20) & 0xFFF
                i += 4
                if size == 0xFFF:
                    size = struct.unpack_from("<I", data, i)[0]; i += 4
                if tag == 67:
                    raw, chars, k = data[i:i + size], [], 0
                    while k + 2 <= len(raw):
                        c = struct.unpack_from("<H", raw, k)[0]
                        if c < 32:
                            if c in (10, 13):
                                chars.append("\n")
                            if c in (1, 2, 3, 9, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
                                if c == 9:
                                    chars.append(" ")
                                k += 16; continue
                            k += 2; continue
                        chars.append(chr(c)); k += 2
                    t = "".join(chars).strip()
                    if t:
                        out.append(t)
                i += size
    finally:
        ole.close()
    if not out:
        raise RuntimeError("글이 없는 HWP(배포용 · 암호 문서일 수 있음)")
    return "\n".join(out)


def read_hwpx(path):
    z = _zip(path)
    out = []
    for n in sorted(z.namelist()):
        if re.match(r"Contents/section\d+\.xml", n):
            out += _xml_text(_zread(z, n).decode("utf-8", "replace"), "hp:p", "hp:t")
    return "\n".join(_unesc(l) for l in out)


def read_txt(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp949", "utf-16"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


READERS = {".docx": read_docx, ".doc": read_doc, ".pptx": read_pptx, ".xlsx": read_xlsx, ".pdf": read_pdf, ".hwp": read_hwp, ".hwpx": read_hwpx,
           ".txt": read_txt, ".md": read_txt, ".csv": read_txt}


def files_in(folder, work=None):
    """폴더(하위 2단까지)의 읽을 수 있는 파일. 작업 폴더 · 임시 파일은 뺌. 이름 순."""
    folder = os.path.abspath(folder)
    work = os.path.abspath(work) if work else None
    res = []
    for root, dirs, fs in os.walk(folder):
        if work and os.path.abspath(root).startswith(work):
            continue
        if root[len(folder):].count(os.sep) >= 2:
            dirs[:] = []
        for f in fs:
            if f.startswith(("~$", ".")) or os.path.splitext(f)[1].lower() not in EXTS or re.search(r"_Draft\.pptx$|_확인목록\.md$", f):
                continue
            p = os.path.join(root, f)
            if os.path.getsize(p) > 80 * 1024 * 1024:
                continue
            res.append(p)
    return sorted(res)


def main(folder, out_md, work=None):
    rows, parts, total = [], [], 0
    fs = files_in(folder, work)
    # 짧고 핵심인 문서(회의록 · RFP · 요청)를 먼저, 길고 일반적인 것(리포트 · 발표)을 뒤에
    key = lambda p: (0 if re.search(r"RFP|요청|회의록|minutes|메모|요구|과업", os.path.basename(p), re.I) else 1, os.path.getsize(p))
    for p in sorted(fs, key=key):
        ext = os.path.splitext(p)[1].lower()
        name = os.path.relpath(p, folder)
        try:
            t = READERS[ext](p)
        except Exception as e:  # 못 읽은 파일은 목록에만 남김
            rows.append({"file": name, "chars": 0, "error": str(e)[:80]})
            continue
        t = re.sub(r"[ \t]+", " ", t)
        t = re.sub(r"\n{3,}", "\n\n", t).strip()
        take = t[:max(0, min(PER_FILE, TOTAL - total))]
        total += len(take)
        rows.append({"file": name, "chars": len(t), "used": len(take)})
        if take:
            parts.append(f"## {name} ({ext[1:]} · 원문 {len(t):,}자{' · 앞부분만' if len(take) < len(t) else ''})\n{take}")
    with open(out_md, "w", encoding="utf-8") as f:
        f.write(f"# 고객 자료 ({folder})\n\n" + "\n\n".join(parts))
    return rows, total
