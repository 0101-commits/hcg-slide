# checkops.js 대조 보고(파이썬 판 check.py vs JS 판)

- 고객 한빛소재 · 공식 수치 2026.04 · 이전 고객명 사전 135건 · 시험 파일 19개
- 실행 A = 기본 입력(고객명·수치 표만). 실행 B = 용어 사전·요청 항목·회의 지시·RFP 핵심어·분량·RFP 주수·공공 규정(무기명 등)까지 모두 켬
- 대조 기준: 항목/수준별 건수, 그리고 (항목·수준·장·칸) 목록 일치
- 결과: 39 / 39 실행이 완전 일치

## work/converted/2b96b804/2b8f906957.pptx (9장)

실행 A: 파이썬 14건 · JS 14건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 4 | 4 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 14/confirm | 6 | 6 | O |

실행 B: 파이썬 56건 · JS 56건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 4 | 4 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 14/confirm | 6 | 6 | O |
| 18/confirm | 1 | 1 | O |
| 19/confirm | 3 | 3 | O |
| 20/fix | 21 | 21 | O |
| 23/confirm | 2 | 2 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 9 | 9 | O |

## work/converted/2b96b804/53c0078373.pptx (62장)

실행 A: 파이썬 130건 · JS 130건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 91 | 91 | O |
| 4/fix | 3 | 3 | O |
| 6/confirm | 1 | 1 | O |
| 9/fix | 2 | 2 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 14 | 14 | O |
| 14/confirm | 6 | 6 | O |
| 15/confirm | 12 | 12 | O |

실행 B: 파이썬 272건 · JS 272건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 91 | 91 | O |
| 4/fix | 3 | 3 | O |
| 6/confirm | 1 | 1 | O |
| 9/fix | 2 | 2 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 14 | 14 | O |
| 14/confirm | 6 | 6 | O |
| 15/confirm | 12 | 12 | O |
| 19/confirm | 1 | 1 | O |
| 20/fix | 68 | 68 | O |
| 24/fix | 2 | 2 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 62 | 62 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/2b96b804/84a498e65e.pptx (18장)

실행 A: 파이썬 14건 · JS 14건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 13/confirm | 1 | 1 | O |
| 14/confirm | 7 | 7 | O |

실행 B: 파이썬 50건 · JS 50건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 13/confirm | 1 | 1 | O |
| 14/confirm | 7 | 7 | O |
| 18/confirm | 3 | 3 | O |
| 19/confirm | 3 | 3 | O |
| 23/confirm | 2 | 2 | O |
| 24/fix | 4 | 4 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 18 | 18 | O |

## work/converted/2b96b804/87b648965c.pptx (65장)

실행 A: 파이썬 76건 · JS 76건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 4 | 4 | O |
| 2/confirm | 1 | 1 | O |
| 6/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 14/confirm | 63 | 63 | O |
| 15/confirm | 1 | 1 | O |

실행 B: 파이썬 233건 · JS 233건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 4 | 4 | O |
| 2/confirm | 1 | 1 | O |
| 6/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 14/confirm | 63 | 63 | O |
| 15/confirm | 1 | 1 | O |
| 19/confirm | 3 | 3 | O |
| 20/fix | 78 | 78 | O |
| 24/fix | 2 | 2 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 65 | 65 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/2b96b804/9013f64d08.pptx (55장)

실행 A: 파이썬 83건 · JS 83건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 6/confirm | 2 | 2 | O |
| 7/confirm | 1 | 1 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 4 | 4 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 13/confirm | 1 | 1 | O |
| 14/confirm | 4 | 4 | O |
| 15/confirm | 4 | 4 | O |
| 16/confirm | 1 | 1 | O |
| 17/confirm | 62 | 62 | O |

실행 B: 파이썬 226건 · JS 226건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 6/confirm | 2 | 2 | O |
| 7/confirm | 1 | 1 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 4 | 4 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 1 | 1 | O |
| 13/confirm | 1 | 1 | O |
| 14/confirm | 4 | 4 | O |
| 15/confirm | 4 | 4 | O |
| 16/confirm | 1 | 1 | O |
| 17/confirm | 62 | 62 | O |
| 18/confirm | 2 | 2 | O |
| 19/confirm | 2 | 2 | O |
| 20/fix | 71 | 71 | O |
| 24/fix | 4 | 4 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 55 | 55 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/2b96b804/a35e6f2492.pptx (18장)

실행 A: 파이썬 6건 · JS 6건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 8/confirm | 1 | 1 | O |
| 9/fix | 2 | 2 | O |
| 11/ref | 1 | 1 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 1 | 1 | O |

실행 B: 파이썬 90건 · JS 90건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 8/confirm | 1 | 1 | O |
| 9/fix | 2 | 2 | O |
| 11/ref | 1 | 1 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 1 | 1 | O |
| 19/confirm | 2 | 2 | O |
| 20/fix | 50 | 50 | O |
| 23/confirm | 1 | 1 | O |
| 24/fix | 7 | 7 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 18 | 18 | O |

## work/converted/2b96b804/a9da42eabf.pptx (58장)

실행 A: 파이썬 15건 · JS 15건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/confirm | 1 | 1 | O |
| 8/confirm | 3 | 3 | O |
| 9/fix | 2 | 2 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 3 | 3 | O |
| 13/confirm | 1 | 1 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 2 | 2 | O |
| 15/confirm | 1 | 1 | O |

실행 B: 파이썬 133건 · JS 133건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/confirm | 1 | 1 | O |
| 8/confirm | 3 | 3 | O |
| 9/fix | 2 | 2 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 3 | 3 | O |
| 13/confirm | 1 | 1 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 2 | 2 | O |
| 15/confirm | 1 | 1 | O |
| 19/confirm | 1 | 1 | O |
| 20/fix | 43 | 43 | O |
| 24/fix | 7 | 7 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 58 | 58 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/2b96b804/c492a723fd.pptx (76장)

실행 A: 파이썬 60건 · JS 60건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 6/confirm | 3 | 3 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 7 | 7 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 4 | 4 | O |
| 13/confirm | 1 | 1 | O |
| 14/confirm | 28 | 28 | O |
| 15/confirm | 12 | 12 | O |
| 22/ref | 2 | 2 | O |

실행 B: 파이썬 492건 · JS 492건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 6/confirm | 3 | 3 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 7 | 7 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 4 | 4 | O |
| 13/confirm | 1 | 1 | O |
| 14/confirm | 28 | 28 | O |
| 15/confirm | 12 | 12 | O |
| 19/confirm | 2 | 2 | O |
| 20/fix | 337 | 337 | O |
| 22/ref | 2 | 2 | O |
| 24/fix | 7 | 7 | O |
| 25/confirm | 1 | 1 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 76 | 76 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/2b96b804/c9189157d2.pptx (18장)

실행 A: 파이썬 33건 · JS 33건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 3 | 3 | O |
| 2/confirm | 1 | 1 | O |
| 2/fix | 5 | 5 | O |
| 7/confirm | 1 | 1 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 5 | 5 | O |
| 14/confirm | 2 | 2 | O |
| 17/confirm | 10 | 10 | O |

실행 B: 파이썬 64건 · JS 64건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 3 | 3 | O |
| 2/confirm | 1 | 1 | O |
| 2/fix | 5 | 5 | O |
| 7/confirm | 1 | 1 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 5 | 5 | O |
| 14/confirm | 2 | 2 | O |
| 17/confirm | 10 | 10 | O |
| 18/confirm | 1 | 1 | O |
| 19/confirm | 3 | 3 | O |
| 23/confirm | 2 | 2 | O |
| 24/fix | 1 | 1 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 18 | 18 | O |

## work/converted/2b96b804/dfbee7fbdd.pptx (10장)

실행 A: 파이썬 14건 · JS 14건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 6/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 2 | 2 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 5 | 5 | O |
| 15/confirm | 4 | 4 | O |

실행 B: 파이썬 53건 · JS 53건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 6/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 2 | 2 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 5 | 5 | O |
| 15/confirm | 4 | 4 | O |
| 18/confirm | 1 | 1 | O |
| 19/confirm | 2 | 2 | O |
| 20/fix | 16 | 16 | O |
| 23/confirm | 1 | 1 | O |
| 24/fix | 3 | 3 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 10 | 10 | O |

## work/converted/2b96b804/e8a5650b28.pptx (50장)

실행 A: 파이썬 132건 · JS 132건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/confirm | 1 | 1 | O |
| 1/fix | 65 | 65 | O |
| 2/confirm | 1 | 1 | O |
| 3/confirm | 5 | 5 | O |
| 3/fix | 1 | 1 | O |
| 4/fix | 19 | 19 | O |
| 6/confirm | 2 | 2 | O |
| 8/confirm | 3 | 3 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 20 | 20 | O |
| 13/confirm | 1 | 1 | O |
| 15/confirm | 5 | 5 | O |
| 16/confirm | 2 | 2 | O |
| 21/confirm | 1 | 1 | O |
| 22/ref | 2 | 2 | O |

실행 B: 파이썬 422건 · JS 422건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/confirm | 1 | 1 | O |
| 1/fix | 65 | 65 | O |
| 2/confirm | 1 | 1 | O |
| 3/confirm | 5 | 5 | O |
| 3/fix | 1 | 1 | O |
| 4/fix | 19 | 19 | O |
| 6/confirm | 2 | 2 | O |
| 8/confirm | 3 | 3 | O |
| 9/fix | 3 | 3 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 20 | 20 | O |
| 13/confirm | 1 | 1 | O |
| 15/confirm | 5 | 5 | O |
| 16/confirm | 2 | 2 | O |
| 19/confirm | 1 | 1 | O |
| 20/fix | 187 | 187 | O |
| 21/confirm | 1 | 1 | O |
| 22/ref | 2 | 2 | O |
| 24/confirm | 8 | 8 | O |
| 24/fix | 34 | 34 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 51 | 51 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/3b39527f/56cc9790f8.pptx (47장)

실행 A: 파이썬 31건 · JS 31건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/confirm | 1 | 1 | O |
| 6/confirm | 1 | 1 | O |
| 9/fix | 6 | 6 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 5 | 5 | O |
| 13/confirm | 1 | 1 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 7 | 7 | O |
| 15/confirm | 8 | 8 | O |

실행 B: 파이썬 136건 · JS 136건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/confirm | 1 | 1 | O |
| 6/confirm | 1 | 1 | O |
| 9/fix | 6 | 6 | O |
| 10/fix | 2 | 2 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 5 | 5 | O |
| 13/confirm | 1 | 1 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 7 | 7 | O |
| 15/confirm | 8 | 8 | O |
| 19/confirm | 2 | 2 | O |
| 20/fix | 33 | 33 | O |
| 24/confirm | 2 | 2 | O |
| 24/fix | 10 | 10 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 47 | 47 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/501bb68e/e66958f3af.pptx (186장)

실행 A: 파이썬 148건 · JS 148건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/fix | 1 | 1 | O |
| 4/fix | 1 | 1 | O |
| 6/confirm | 9 | 9 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 5 | 5 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 40 | 40 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 47 | 47 | O |
| 15/confirm | 36 | 36 | O |
| 17/confirm | 3 | 3 | O |
| 22/ref | 2 | 2 | O |

실행 B: 파이썬 835건 · JS 835건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 2/fix | 1 | 1 | O |
| 4/fix | 1 | 1 | O |
| 6/confirm | 9 | 9 | O |
| 7/fix | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 5 | 5 | O |
| 11/confirm | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 40 | 40 | O |
| 13/ref | 1 | 1 | O |
| 14/confirm | 47 | 47 | O |
| 15/confirm | 36 | 36 | O |
| 17/confirm | 3 | 3 | O |
| 19/confirm | 1 | 1 | O |
| 20/fix | 480 | 480 | O |
| 22/ref | 2 | 2 | O |
| 24/fix | 10 | 10 | O |
| 25/confirm | 1 | 1 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 186 | 186 | O |
| 28/confirm | 2 | 2 | O |

## work/converted/617d4ce7/7d2068d894.pptx (12장)

실행 A: 파이썬 14건 · JS 14건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 2 | 2 | O |
| 4/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 4 | 4 | O |
| 15/confirm | 2 | 2 | O |

실행 B: 파이썬 74건 · JS 74건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 2 | 2 | O |
| 4/confirm | 1 | 1 | O |
| 8/confirm | 1 | 1 | O |
| 9/fix | 3 | 3 | O |
| 10/fix | 6 | 6 | O |
| 11/ref | 1 | 1 | O |
| 12/confirm | 4 | 4 | O |
| 15/confirm | 2 | 2 | O |
| 19/confirm | 1 | 1 | O |
| 20/fix | 31 | 31 | O |
| 24/fix | 4 | 4 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 12 | 12 | O |

## work/spike_struct/s1_card/orig.pptx (1장)

실행 A: 파이썬 3건 · JS 3건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 15/confirm | 1 | 1 | O |

실행 B: 파이썬 23건 · JS 23건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 15/confirm | 1 | 1 | O |
| 18/confirm | 3 | 3 | O |
| 19/confirm | 3 | 3 | O |
| 20/fix | 2 | 2 | O |
| 23/confirm | 2 | 2 | O |
| 24/fix | 4 | 4 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |

## work/spike_struct/s2_asis/orig.pptx (1장)

실행 A: 파이썬 2건 · JS 2건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 9/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |

실행 B: 파이썬 25건 · JS 25건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 9/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 18/confirm | 3 | 3 | O |
| 19/confirm | 3 | 3 | O |
| 20/fix | 5 | 5 | O |
| 23/confirm | 2 | 2 | O |
| 24/fix | 3 | 3 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |
| 27/fix | 1 | 1 | O |

## work/spike_struct/s3_option/orig.pptx (1장)

실행 A: 파이썬 3건 · JS 3건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 22/fix | 1 | 1 | O |

실행 B: 파이썬 23건 · JS 23건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 18/confirm | 3 | 3 | O |
| 19/confirm | 3 | 3 | O |
| 20/fix | 1 | 1 | O |
| 22/fix | 1 | 1 | O |
| 23/confirm | 2 | 2 | O |
| 24/fix | 5 | 5 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |

## work/spike_struct/s4_gantt/orig.pptx (1장)

실행 A: 파이썬 1건 · JS 1건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 11/ref | 1 | 1 | O |

실행 B: 파이썬 18건 · JS 18건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 11/ref | 1 | 1 | O |
| 18/confirm | 3 | 3 | O |
| 19/confirm | 3 | 3 | O |
| 23/confirm | 1 | 1 | O |
| 24/fix | 4 | 4 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |

## work/spike_struct/s5_org/orig.pptx (1장)

실행 A: 파이썬 3건 · JS 3건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 15/confirm | 1 | 1 | O |

실행 B: 파이썬 20건 · JS 20건 → 전부 일치

| 항목/수준 | 파이썬 | JS | 일치 |
|---|---|---|---|
| 1/fix | 1 | 1 | O |
| 11/ref | 1 | 1 | O |
| 15/confirm | 1 | 1 | O |
| 18/confirm | 3 | 3 | O |
| 19/confirm | 3 | 3 | O |
| 23/confirm | 1 | 1 | O |
| 24/fix | 4 | 4 | O |
| 25/fix | 1 | 1 | O |
| 26/fix | 5 | 5 | O |

## 합성 시험(고객명 묶기·기준월·습니다체)

실행: 파이썬 67건 · JS 67건 → 전부 일치

- 고객명 묶음(확인할 것 「고객명 N개」) 4건 · 4개 이하 장표의 이름별 고칠 것 16건 · 태그 이름이 든 장표의 고칠 것 1건
- 기준월 붙은 「고객 950사」 고칠 것 4건 · 기준월 없는 「고객 820사」 확인할 것 4건
- 습니다체 확인할 것 12건(인사말·인용 포함 문장이 잡힌 건수 0)
