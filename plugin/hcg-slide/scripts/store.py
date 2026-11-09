# 창고(Library) 받기 · 개인 열쇠 보관.
# 열쇠: /key 화면에서 받은 HCG_열쇠.hcgkey 를 Windows 사용자 잠금(DPAPI)으로 이 PC · 이 사용자만 풀 수 있게 저장하고 원본은 지움.
#        열쇠 글자는 화면 · 대화에 찍지 않음.
# 받은 파일은 <데이터 폴더>/cache 에 두고, 서버 판(version)이 바뀔 때만 다시 받음.
import ctypes, ctypes.wintypes as wt, glob, json, os, sys, time, urllib.request, urllib.error

SERVER = "https://hcg-slides.pages.dev"


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, protect: bool) -> bytes:
    src = _Blob(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(src), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("Windows 사용자 잠금 실패")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


def downloads_dir():
    # 「다운로드」 폴더(옮겨 둔 경우 포함) — Shell 알려진 폴더 FOLDERID_Downloads
    try:
        from ctypes import windll, c_wchar_p, byref
        import uuid
        fid = uuid.UUID("{374DE290-123F-4565-9164-39C4925E467B}")
        guid = (ctypes.c_byte * 16).from_buffer_copy(fid.bytes_le)
        p = c_wchar_p()
        if windll.shell32.SHGetKnownFolderPath(byref(guid), 0, None, byref(p)) == 0:
            path = p.value
            windll.ole32.CoTaskMemFree(p)
            return path
    except Exception:
        pass
    return os.path.join(os.path.expanduser("~"), "Downloads")


def key_path(data):
    return os.path.join(data, "key.dat")


def import_key(data):
    """다운로드 폴더의 가장 최근 *.hcgkey 를 잠가 저장하고 원본을 지움. 열쇠 글자는 돌려주지 않음."""
    files = sorted(glob.glob(os.path.join(downloads_dir(), "*.hcgkey")), key=os.path.getmtime)
    if not files:
        return {"ok": False, "why": "다운로드 폴더에 .hcgkey 파일이 없음 — https://hcg-slides.pages.dev/key 에서 「열쇠 받기」"}
    src = files[-1]
    info = json.load(open(src, encoding="utf-8"))
    os.makedirs(data, exist_ok=True)
    blob = _dpapi(json.dumps({"key": info["key"], "email": info.get("email"), "server": info.get("server") or SERVER}).encode("utf-8"), True)
    open(key_path(data), "wb").write(blob)
    for f in files:  # 남은 열쇠 파일은 모두 지움
        try:
            os.remove(f)
        except OSError:
            pass
    return {"ok": True, "email": info.get("email"), "removed": len(files)}


def _key(data):
    p = key_path(data)
    if not os.path.exists(p):
        return None
    return json.loads(_dpapi(open(p, "rb").read(), False).decode("utf-8"))


def get(data, route, timeout=60):
    k = _key(data)
    if not k:
        raise SystemExit("열쇠 없음 — /hcg-draft 의 열쇠 준비 단계를 먼저")
    req = urllib.request.Request(f"{k['server'] or SERVER}/kapi/{route}", headers={"authorization": f"Bearer {k['key']}", "accept-encoding": "identity", "user-agent": "hcg-slide-plugin"})
    try:
        return urllib.request.urlopen(req, timeout=timeout).read()
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise SystemExit("열쇠가 거부됨(지워졌거나 잘못됨) — /key 에서 새로 받고 다시 준비")
        raise


def cache_dir(data):
    d = os.path.join(data, "cache")
    os.makedirs(os.path.join(d, "pptx"), exist_ok=True)
    return d


def sync(data):
    """목록 · 그림 목록 · 유형표 · 틀을 서버 판이 바뀌었을 때만 받음."""
    c = cache_dir(data)
    ver = json.loads(get(data, "version"))
    have = json.load(open(os.path.join(c, "version.json"), encoding="utf-8")) if os.path.exists(os.path.join(c, "version.json")) else {}
    got = []
    for name, route, fn in [("meta", "meta", "meta.json"), ("media", "media", "media_index.json"), ("types", "types", "types.json"), ("frame", "frame/std.pptx", "frame_std.pptx")]:
        if ver.get(name) and (have.get(name) != ver[name] or not os.path.exists(os.path.join(c, fn))):
            open(os.path.join(c, fn), "wb").write(get(data, route, timeout=180))
            have[name] = ver[name]
            got.append(fn)
    json.dump(have, open(os.path.join(c, "version.json"), "w", encoding="utf-8"))
    return got


def slide(data, sid):
    p = os.path.join(cache_dir(data), "pptx", f"{sid}.pptx")
    if not os.path.exists(p):
        open(p, "wb").write(get(data, f"slide/{sid}.pptx"))
    return p


def status(data):
    k = None
    try:
        k = _key(data)
    except Exception as e:
        return {"key": False, "why": f"열쇠를 못 풂: {e}"}
    if not k:
        return {"key": False}
    try:
        me = json.loads(get(data, "me", timeout=20))
        return {"key": True, "email": me.get("email"), "server": True}
    except SystemExit as e:
        return {"key": True, "server": False, "why": str(e)}
    except Exception as e:
        return {"key": True, "server": False, "why": f"서버에 못 닿음: {e}"}


if __name__ == "__main__":
    data = sys.argv[2]
    cmd = sys.argv[1]
    if cmd == "import":
        print(json.dumps(import_key(data), ensure_ascii=False))
    elif cmd == "status":
        print(json.dumps(status(data), ensure_ascii=False))
    elif cmd == "sync":
        t = time.time(); print(json.dumps({"got": sync(data), "sec": round(time.time() - t, 1)}, ensure_ascii=False))
