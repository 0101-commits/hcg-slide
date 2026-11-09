-- 11-04 판 2.0: 유형(AI 분류) · 장(제안서 흐름) · 확신
ALTER TABLE slides ADD COLUMN type TEXT;
ALTER TABLE slides ADD COLUMN chap TEXT;
ALTER TABLE slides ADD COLUMN tconf TEXT;
CREATE INDEX IF NOT EXISTS slides_chap ON slides(chap);
