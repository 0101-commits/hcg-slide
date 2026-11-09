# Item 미리보기가 거의 빈 것(글자만 남은 도형 등) 빼기 — 유형마다 기준(괄호 · 띠는 선이 가늘어 낮게)
import json
from pathlib import Path
from PIL import Image
P = Path(__file__).resolve().parent.parent / "pane" / "public" / "pane"
TH = {"brace": 0.002, "bar": 0.004, "box": 0.006}
kinds = json.load(open(P / "items.json", encoding="utf-8"))
for k in kinds:
    keep = []
    for it in k["items"]:
        f = P / "items" / f"{it['id']}.png"
        im = Image.open(f).convert("L")
        ink = sum(1 for v in im.tobytes() if v < 235) / (im.width * im.height)
        if ink < TH.get(k["code"], 0.012):
            f.unlink()
            continue
        keep.append(it)
    k["items"] = keep
    print(k["code"], len(keep))
json.dump(kinds, open(P / "items.json", "w", encoding="utf-8"), ensure_ascii=False)
