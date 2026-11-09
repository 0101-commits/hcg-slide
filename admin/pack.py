# 설치본 묶기 — app/dist/HCG장표_설치.zip
# 실행: ..\..\.venv\Scripts\python -X utf8 pack.py [--url https://로컬·시험 주소]
# 안: 설치.bat · 제거.bat · 읽어주세요.txt · HCG장표_사용설명서.html · addin\(manifest.xml · install.ps1 · uninstall.ps1)
#     skills\(Claude Code 용 폴더 3개) · Claude앱_스킬\(Claude 앱 · PowerPoint 에 올릴 zip 3개)
import argparse, json, shutil, zipfile
from pathlib import Path

APP = Path(__file__).resolve().parent.parent
DIST = APP / "dist"
SKIP = {"__pycache__", "_test", ".pytest_cache"}

README = """HCG 장표 설치 안내 (판 {ver})

쓰는 법은 같은 폴더의 「HCG장표_사용설명서.html」을 두 번 클릭해 보세요(브라우저로 열림).

1. 이 zip 을 아무 폴더에나 풉니다.
2. 「설치.bat」를 두 번 클릭합니다.
   처음 한 번 Windows 가 「알 수 없는 게시자」 경고를 띄우면 「추가 정보 → 실행」을 누릅니다.
   「설치 끝」 창이 뜨면 됩니다. 관리자 권한은 필요 없습니다.
3. PowerPoint 를 모두 닫고 다시 켭니다. 처음 한 번만 추가 기능을 넣습니다.
   홈 → 추가 기능 → 맨 아래 「추가 기능 더 보기」 → 위쪽 「공유 폴더」 → 「HCG 장표」 → 「추가」.
   그다음부터는 PowerPoint 를 켤 때마다 홈 리본 오른쪽에 「HCG 장표」 단추가 그대로 있습니다.
   (「설치 끝」 창이 「개발자 추가 기능 방식」이라고 하면 이 PC 는 공유 폴더 방식을 못 씁니다.
    그때는 PowerPoint 를 켤 때마다 홈 → 추가 기능 → 「개발자 추가 기능」의 「HCG 장표」를 누릅니다.)
   단추를 누르고 처음 한 번 회사 메일로 로그인합니다(메일로 온 숫자 입력).

4. Claude 스킬 올리기(한 번, 판이 바뀌면 다시)
   Claude 앱(데스크톱 · 웹) → 설정 → 기능(또는 Customize) → 스킬 → 「스킬 올리기」에서
   「Claude앱_스킬」 폴더의 zip 3개를 하나씩 올리고 켭니다.
     hcg-draft.zip        초안 — 요청서 · 회의 메모 → 목차 · 장별 주장 문장
     hcg-restructure.zip  구조 바꾸기 — 칸은 두고 글만, 열 · 행 · 간트 막대
     hcg-check.zip        점검 — 이전 고객명 · 자리표시자 · 회사 수치 등 28항목
   앱에 올린 스킬은 PowerPoint 의 Claude 작업창에서도 그대로 보입니다.
   Claude Code(CLI)를 쓰는 사람은 설치.bat 가 스킬을 %USERPROFILE%\\.claude\\skills 에 복사해 둡니다.

5. 공식 「Claude」 단추가 PowerPoint 에 없으면
   PowerPoint → 홈 → 추가 기능 → 「Claude」 검색 → 추가 → 회사 Claude 계정으로 로그인.

지우기: 「제거.bat」를 두 번 클릭합니다.
문제가 생기면 %LOCALAPPDATA%\\HCG\\install.log 와 화면을 관리자에게 보내 주세요.
"""

BAT_IN = '@echo off\r\npowershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0addin\\install.ps1"\r\n'
BAT_OUT = '@echo off\r\npowershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0addin\\uninstall.ps1"\r\n'


def render_manifest(cfg, url):
    t = (APP / "addin" / "manifest.template.xml").read_text(encoding="utf-8")
    acc = f"    <AppDomain>https://{cfg['access_team']}.cloudflareaccess.com</AppDomain>" if cfg.get("access_team") else ""
    out = (t.replace("{{ADDIN_ID}}", cfg["addin_id"]).replace("{{VERSION}}", cfg["version"])
           .replace("{{URL}}", url.rstrip("/")).replace("{{ACCESS_DOMAIN}}", acc))
    assert "{{" not in out, "채우지 못한 칸이 남음"
    return out


def copy_skill(src, dst):
    shutil.copytree(src, dst, ignore=lambda d, names: [n for n in names if n in SKIP or n.endswith(".pyc")])


def zip_dir(folder, zpath, arc_root):
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(folder.rglob("*")):
            if p.is_file():
                z.write(p, arc_root + "/" + p.relative_to(folder).as_posix())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", help="작업창 주소(기본 config.json pages_url)")
    ap.add_argument("--out", default="HCG장표_설치.zip")
    a = ap.parse_args()
    cfg = json.loads((APP / "config.json").read_text(encoding="utf-8"))
    url = a.url or cfg["pages_url"]

    stage = DIST / "stage"
    if stage.exists():
        shutil.rmtree(stage)
    (stage / "addin").mkdir(parents=True)
    (stage / "addin" / "manifest.xml").write_text(render_manifest(cfg, url), encoding="utf-8")
    for fn in ("install.ps1", "uninstall.ps1"):  # Windows PowerShell 5.1 은 BOM 이 있어야 한글을 읽음
        (stage / "addin" / fn).write_text((APP / "install" / fn).read_text(encoding="utf-8"), encoding="utf-8-sig")
    (stage / "설치.bat").write_bytes(BAT_IN.encode("ascii"))
    (stage / "제거.bat").write_bytes(BAT_OUT.encode("ascii"))
    man = APP / "docs" / "HCG장표_사용설명서.html"  # 16:9 장표형 설명서(그림까지 한 파일) — docs/build_manual.py 로 만듦
    if man.exists():
        shutil.copy2(man, stage / man.name)
    else:
        print("주의: 사용설명서 없음 —", man)
    (stage / "읽어주세요.txt").write_text(README.format(ver=cfg["version"]).replace("\n", "\r\n"), encoding="utf-8-sig")

    skills = sorted(p for p in (APP / "skills").iterdir() if p.is_dir() and p.name not in SKIP and (p / "SKILL.md").exists())
    (stage / "skills").mkdir()
    (stage / "Claude앱_스킬").mkdir()
    for s in skills:
        copy_skill(s, stage / "skills" / s.name)
        zip_dir(stage / "skills" / s.name, stage / "Claude앱_스킬" / f"{s.name}.zip", s.name)

    DIST.mkdir(exist_ok=True)
    zout = DIST / a.out
    with zipfile.ZipFile(zout, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(stage.rglob("*")):
            if p.is_file():
                z.write(p, p.relative_to(stage).as_posix())
    print(zout, zout.stat().st_size // 1024, "KB · 스킬", [s.name for s in skills], "· 주소", url)


if __name__ == "__main__":
    main()
