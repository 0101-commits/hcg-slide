# 점검 탭이 쓰는 사전을 스킬(hcg-check)에서 작업창으로 복사 — 고객명 사전 · 회사 수치 표를 고친 뒤 한 번 돌리고 deploy.mjs pages
# 실행: ..\..\.venv\Scripts\python -X utf8 sync_check_refs.py
import json
from pathlib import Path

APP = Path(__file__).resolve().parent.parent
REF = APP / "skills" / "hcg-check" / "references"
OUT = APP / "pane" / "public" / "pane" / "check"
OUT.mkdir(parents=True, exist_ok=True)
names = [l.strip() for l in open(REF / "client_names.txt", encoding="utf-8") if l.strip() and not l.startswith("#")]
(OUT / "client_names.json").write_text(json.dumps(names, ensure_ascii=False), encoding="utf-8")
(OUT / "company_figures.json").write_text((REF / "company_figures.json").read_text(encoding="utf-8"), encoding="utf-8")
print("고객명", len(names), "· 회사 수치 표 복사 →", OUT)
