"""manual_slides.src.html 의 img/*.png 를 base64 로 끼워 HCG장표_사용설명서.html 을 만든다."""
import base64
import re
from pathlib import Path

here = Path(__file__).resolve().parent
src = (here / "manual_slides.src.html").read_text(encoding="utf-8")


def inline(m):
    p = here / "img" / m.group(1)
    data = base64.b64encode(p.read_bytes()).decode("ascii")
    return 'src="data:image/png;base64,' + data + '"'


out = re.sub(r'src="img/([^"]+\.png)"', inline, src)
dest = here / "HCG장표_사용설명서.html"
dest.write_text(out, encoding="utf-8")
print(dest, round(dest.stat().st_size / 1024), "KB")
