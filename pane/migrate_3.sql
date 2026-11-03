-- 5단계: 단계 칸(1 = 상 · 별표 · 표준, 2 = 중)
ALTER TABLE slides ADD COLUMN tier INTEGER;
CREATE INDEX IF NOT EXISTS slides_tier ON slides(tier);
