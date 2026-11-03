-- 「HCG 장표」 창고 검색 목록
CREATE TABLE IF NOT EXISTS slides (
  id TEXT PRIMARY KEY,        -- s00197 (지난 장표 분류 번호 k)
  k INTEGER NOT NULL,
  title TEXT, text TEXT,      -- 장표 제목 · 글 전문(검색용)
  olds TEXT,                  -- |이전 고객명|...| (글에서 찾은 것)
  mods TEXT, mod TEXT,        -- |HR 모듈|...| · 첫 모듈
  p TEXT, p_name TEXT,        -- 목적 코드 · 이름
  s TEXT, s_name TEXT,        -- 구조 코드 · 이름
  grade TEXT, star INTEGER,   -- 재사용 등급 상·중·하 · 별표
  kind TEXT, kind_name TEXT,  -- 자주 쓰는 표준 장표 6종
  action TEXT,                -- 넣은 뒤 할 일(그대로 · 고객명만 · 칸 글 · 숫자 · 차트 · 막대 · 주차)
  client TEXT, year TEXT,     -- 원본 고객 · 연도
  industry TEXT,              -- 고객 업종(2단계)
  doc TEXT, page INTEGER,     -- 원본 파일 이름 · 쪽
  bytes INTEGER,
  added TEXT DEFAULT (date('now')),
  result TEXT                 -- 결과 기록(선택): 수주 · drop
);
CREATE INDEX IF NOT EXISTS slides_p ON slides(p);
CREATE INDEX IF NOT EXISTS slides_s ON slides(s);
