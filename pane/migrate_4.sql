-- 11-04: 장표 이름(어깨끈 마지막 칸) · 묶음(같은 이름) · 모양(거의 같은 장) · 자주 쓰는 표준 장표
ALTER TABLE slides ADD COLUMN name TEXT;
ALTER TABLE slides ADD COLUMN head TEXT;
ALTER TABLE slides ADD COLUMN fam TEXT;
ALTER TABLE slides ADD COLUMN fam_name TEXT;
ALTER TABLE slides ADD COLUMN fam_n INTEGER;
ALTER TABLE slides ADD COLUMN fam_v INTEGER;
ALTER TABLE slides ADD COLUMN fam_docs INTEGER;
ALTER TABLE slides ADD COLUMN var TEXT;
ALTER TABLE slides ADD COLUMN var_n INTEGER;
ALTER TABLE slides ADD COLUMN var_rep INTEGER;
ALTER TABLE slides ADD COLUMN var_ord INTEGER;
ALTER TABLE slides ADD COLUMN std INTEGER;
CREATE INDEX IF NOT EXISTS slides_fam ON slides(fam);
CREATE INDEX IF NOT EXISTS slides_std ON slides(std);
