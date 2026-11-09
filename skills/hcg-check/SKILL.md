---
name: hcg-check
description: HCG 제안서 점검 28항목(이전 고객명·자리표시자·회사 수치·일정 주수·무기명 등)을 보고서로 만듦. 「점검」「검수」「고객명 잔재」「/hcg-점검」 요청에 사용.
---

# HCG 제안서 점검 (hcg-check)

완성했거나 작성 중인 제안서(PPTX)를 28가지로 점검해 「고칠 것 / 확인할 것 / 참고」 보고서를 만든다. 사람들이 「/hcg-점검」이라 부르던 작업임. 지난 제안 39건에서 사람이 손으로 고친 것과 공공·입찰형 규정을 합쳐 28가지를 뽑았고, 자동으로 찾는 것(16)과 후보만 뽑고 문장은 Claude 가 판단하는 것(12)으로 나뉜다.
문체는 우리말 「~함」. 파일·함수 이름만 영문.

## 28항목 한눈에

| 무리 | 자동 | 반자동 |
|---|---|---|
| 글자 규칙 | 1 이전 고객명 잔재 · 2 자리표시자·작업 메모 · 3 회사 소개 수치 · 4 예시·가설 표시 · 5 익명 표기 통일 | 6 출처·기준일 |
| 장표 구조 | 7 제목 번호·접두 · 8 목차 간지 · 9 레이아웃 이름 · 10 일정 숫자 동기화 | 11 분량 · 12 도식 많음·없음 · 13 [참고] 분리 |
| 문장 판단 | 17 종결 어미 | 14 결론형 핵심 문장 · 15 단정·방어 유발 문구 · 16 Why HCG 주장형 |
| 입력 대조 | 20 고객 표기 · 22 가격 분리 | 18 요청 항목 반영 · 19 RFP 핵심어 · 21 투입 인력 주요 고객 · 23 마지막 회의 지시 |
| 공공·입찰형 | 24 무기명 위반 · 25 수행 실적 기간 · 26 지정 목차·필수 장표 · 27 지정 표지·서식 | 28 쪽수 |

항목마다 찾는 법·고치는 법·근거 건수는 `references/items.md`. 공공·입찰형 5항목(24~28)은 발주처 규정(`rules.json`)을 받았을 때만 켠다.

## 어디서 돌리느냐

### 1. Claude 앱(데스크톱·웹)
사용자가 PPTX 를 올림 → 코드 실행 공간에서 `scripts/check.py` 를 돌려 `report.md` 를 만들고, 반자동 후보를 읽어 판단을 덧붙여 돌려준다. 점검만 하고 파일은 고치지 않는다(고쳐 달라고 하면 `hcg-restructure` 로 새 파일을 만든다). 필요한 것: python-pptx(없으면 `pip install python-pptx`).

### 2. Claude for PowerPoint
파일 시스템·파이썬 없음 → 스크립트 대신 같은 항목을 지침으로 점검한다. 먼저 `references/hcg_marks.md` 의 코드로 「HCG 장표」 작업창이 남긴 표시(새 고객명 `HCG_CLIENT`, 넣은 장의 원본 고객명 `HCG_OLD` 등)를 읽는다. `read_slide_text`·`list_slide_shapes` 로 장 글·도형·레이아웃을 읽고, `references/items.md` 의 「찾는 법」 칼럼대로 직접 찾는다. 사전은 `references/client_names.txt`·`company_figures.json` 을 읽어 쓴다. 고칠 것은 사용자가 요청하면 `edit_slide_text` 로 칸 글만 바꾸고 `verify_slide_visual` 로 넘침을 확인한다. 보고서는 대화에 표로 적는다(형식은 `references/input_schemas.md` 끝).

### 3. Claude Code(CLI)
PC 의 파이썬 + python-pptx 로 `scripts/` 를 직접 돌린다. 폴더째 일괄 점검 가능(파일마다 `check.py` 를 돌려 보고서를 모으고 요약표를 만든다). 원본 제안서 폴더(`OneDrive … 02 Proposal` 등)는 읽기만 하고, 보고서는 다른 폴더에 쓴다. 실행 파이썬은 환경에 맞게(한글 출력이 깨지면 `python -X utf8`).

## 절차

1. 점검할 파일과 맥락을 확인한다: 새 고객명(`--client`), 문서 유형, 공공·입찰형 여부와 발주처 규정, 요청 항목, RFP 핵심어·기간, 분량 지시, 회의 지시. 모르는 것은 묻되, 없어도 가능한 항목은 먼저 돌린다.
2. 입력 파일을 만든다(`references/input_schemas.md`): 공공·입찰형이면 `rules.json`, 필요하면 `terms.json`·`requests.json`·`meeting.json`.
3. 실행:
   ```
   python scripts/check.py <pptx> --out report.md --client 새고객명 \
     --old-names references/client_names.txt --figures references/company_figures.json \
     [--public rules.json] [--terms terms.json] [--requests requests.json] \
     [--rfp-keywords "a,b"] [--meeting meeting.json] [--pages 40] [--rfp-weeks 8-12]
   ```
   글만 뽑아 보고 싶으면 `python scripts/extract_text.py <pptx> --out text.md`(또는 `.json`). Claude 는 이것으로 모든 장의 글을 읽고 반자동 항목의 문장을 판단한다.
4. 보고서를 읽고 정리해 돌려준다:
   - 「고칠 것」: 그대로 전달하고, 고칠 안(바꿀 말)을 구체적으로 적는다.
   - 「확인할 것」: 후보를 장표 내용과 함께 읽고 진짜 문제인지 판단한다. 14·15·16 은 고칠 문장 후보 2개를 쓴다(`../hcg-draft/references/hcg_style.md` 문체). 18 은 요청 대응표의 판정(충분/부분/누락)을 확정한다.
   - 「참고」: 장수 등 알려 두는 것.
   - 우선순위: 이전 고객명 · 회사 수치 · 자리표시자 · 일정 주수 · 가격 · (무기명이면) 무기명 위반을 먼저.
5. 고치기를 원하면 `hcg-restructure` 의 `apply.py`(고객명 바꾸기·칸 글 교체)로 새 파일을 만들고 다시 점검한다. 원본은 덮어쓰지 않는다.

## 보고서 형식

「고칠 것 / 확인할 것 / 참고」 세 표. 열은 항목 · 장 · 칸 · 내용 · 고치는 법. 끝에 28항목 처리 현황(실행 / 후보 / 입력 없음 / 꺼짐)이 붙는다. 장 번호와 칸 이름이 들어간다.

## references 읽는 때

| 파일 | 언제 |
|---|---|
| `references/items.md` | 28항목의 정의·찾는 법·고치는 법. PowerPoint 환경에서 직접 점검할 때, 반자동 후보를 판단할 때 |
| `references/patterns.md` | 후보를 해석하거나 고칠 문장 예(전 → 후)·회사 수치 변천·단정 문구 완화 예가 필요할 때 |
| `references/input_schemas.md` | rules·terms·requests·meeting 파일을 만들 때, 옵션 목록 |
| `references/client_names.txt` | 이전 고객명 사전(1번). 관리자가 `admin/make_client_names.py` 로 갱신 |
| `references/company_figures.json` | 회사 수치 표(3번). 공식 값 = 2026.04 기준 958사 · 219명 |
| `references/hcg_marks.md` | 「HCG 장표」 작업창이 파일에 남긴 표시(새 고객명 · 넣은 장 출처 · 원본 고객명 · 칸 번호 상자)를 읽을 때. PowerPoint 에서 점검을 시작할 때 한 번 |

## 지킬 것

- 회사 수치 공식 값은 2026.04 기준 고객 958사 · HR 인력 219명(관리자 확정 2026-11-03). 다른 값(950개 · 250명, 875사 등)은 옛 기준으로 보고 이 값으로 바꾸게 한다. 사용자가 더 새 값을 알려 주면 그 값을 따른다.
- 자동 항목도 문맥 오탐이 있다: 사례·고객사 장표의 이전 고객명, 무기명 인력 표의 OOO, 일정 옵션 2안의 서로 다른 주수. 「고칠 것」을 그대로 믿고 일괄 수정하지 말고 문맥을 확인한다.
- 텍스트 추출에 안 잡히는 것(그림으로 된 로고, 차트·SmartArt 안의 글)이 있다. 무기명 점검에서는 이 한계를 사용자에게 반드시 알리고 눈으로 확인하게 한다.
- 점검 결과는 「HCG 가 실제로 한 방식」과의 차이일 뿐이고 수주 가능성을 말하는 것이 아니다.
- 입력이 없어 건너뛴 항목은 보고서의 「점검 현황」에 그대로 둔다. 건너뛴 것을 통과로 말하지 않는다.
