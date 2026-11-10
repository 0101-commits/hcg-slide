# 제안서 Draft 실행 묶음 만들기 — 스킬 skills/hcg-draft 가 처음 실행 때 GitHub 릴리스에서 받아 가는 zip
# 실행: ..\..\.venv\Scripts\python -X utf8 pack_draft.py --ver 0.6.0 [--release]
# 안: plugin/hcg-slide 의 bin · python(임베디드 3.10 + python-pptx 등) · scripts · assets
# 하는 일: dist/hcg-draft-runtime-<판>.zip 만들기 → SHA-256 → skills/hcg-draft/runtime.json 고쳐 쓰기
#          --release 면 gh 로 릴리스 draft-<판> 를 만들고 zip 을 올림(저장소가 공개여야 팀원이 받음)
import argparse, hashlib, json, subprocess, zipfile
from pathlib import Path

APP = Path(__file__).resolve().parent.parent
SRC = APP / "plugin" / "hcg-slide"
PARTS = ("bin", "python", "scripts", "assets")
SKIP = {"__pycache__", ".omc", ".pytest_cache"}
REPO = "0101-commits/hcg-slide"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ver", required=True)
    ap.add_argument("--release", action="store_true")
    a = ap.parse_args()
    assert (SRC / "python" / "python.exe").exists(), f"Python 실행기 없음: {SRC / 'python'}"

    out = APP / "dist" / f"hcg-draft-runtime-{a.ver}.zip"
    out.parent.mkdir(exist_ok=True)
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for part in PARTS:
            for p in sorted((SRC / part).rglob("*")):
                if not p.is_file() or p.suffix == ".pyc" or SKIP & set(p.relative_to(SRC).parts):
                    continue
                zi = zipfile.ZipInfo(p.relative_to(SRC).as_posix(), date_time=(2026, 1, 1, 0, 0, 0))  # 같은 내용이면 같은 zip
                zi.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(zi, p.read_bytes())
                n += 1
    sha = hashlib.sha256(out.read_bytes()).hexdigest()
    tag = f"draft-{a.ver}"
    url = f"https://github.com/{REPO}/releases/download/{tag}/{out.name}"
    rj = APP / "skills" / "hcg-draft" / "runtime.json"
    rj.write_text(json.dumps({"version": a.ver, "url": url, "sha256": sha}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"{out.name} · 파일 {n}개 · {out.stat().st_size / 1048576:.1f}MB · sha256 {sha[:12]}… → {rj.relative_to(APP)}")

    if a.release:
        have = subprocess.run(["gh", "release", "view", tag, "-R", REPO], capture_output=True).returncode == 0
        if have:
            subprocess.run(["gh", "release", "upload", tag, str(out), "--clobber", "-R", REPO], check=True)
        else:
            subprocess.run(["gh", "release", "create", tag, str(out), "-R", REPO, "--latest=false",
                            "--title", f"제안서 Draft 실행 묶음 {a.ver}",
                            "--notes", "스킬 hcg-draft 가 처음 실행 때 저절로 받아 가는 실행 묶음(Python · 단계 스크립트). 직접 받을 필요 없음."], check=True)
        print("올림:", url)


if __name__ == "__main__":
    main()
