-- 2단계: 업종 칸 추가(이미 만든 표에 한 번만)
ALTER TABLE slides ADD COLUMN industry TEXT;
CREATE INDEX IF NOT EXISTS slides_ind ON slides(industry);
