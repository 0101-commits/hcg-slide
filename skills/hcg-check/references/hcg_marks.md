# 「HCG 장표」 작업창이 파일에 남기는 표시

「HCG 장표」 작업창과 Claude 는 서로 직접 부를 수 없다. 대신 작업창이 **같은 파일에 표시(태그)** 를 남기고, Claude 가 그것을 읽는다. 작업창의 「Claude에게 ›」 단추는 요청 문장을 복사해 주고, 사용자가 Claude 작업창에 붙여넣는다.

## 표시 목록

| 자리 | 이름 | 뜻 |
|---|---|---|
| 파일 전체 | `HCG_CLIENT` | 작업창 「새 고객명」. 있으면 새 고객명으로 그대로 쓰고 다시 묻지 않는다 |
| 장 | `HCG_SRC` | Library 장표 번호(Library · Component 로 삽입한 장) |
| 장 | `HCG_DOC` | 원본 제안서 이름과 쪽 |
| 장 | `HCG_ADDED` | 넣은 날(YYYY-MM-DD) |
| 장 | `HCG_OLD` | 원본 고객명. 여러 개면 `a\|b` |
| 장 | `HCG_REPLACED` | 넣을 때 바꾼 말. `옛→새\|옛→새` |
| 장 | `HCG_PART` | Component 코드(간트 · 옵션 비교표 · 조직도 등 빈 틀, 글 자리 ///) |
| 장 | `HCG_ITEM` | 끼워 넣은 Item(도형 하나 · 작은 묶음, 글 자리 ///) |
| 장 | `HCG_PLAN_TYPE` · `HCG_PLAN_HEAD` · `HCG_PLAN_MSG` · `HCG_PLAN_MOD` | Dummy 장 — 유형 코드(T101~T604) · 어깨끈 · 핵심 문장 · HR 모듈 |
| 장 | `HCG_EDIT` · `HCG_EDITED` | 「칸」 탭에서 열 · 행 수를 바꾼 장(종류 · 날짜) |
| 도형 이름 | `HCG_NUM_n` | 「칸 번호 붙이기」로 얹은 번호 상자. 내용이 아니므로 점검 · 고치기에서 건너뛴다. 제출 전에는 지워야 한다(작업창 「칸 번호 떼기」) |

## 읽고 나서 할 일

- `HCG_OLD` 이름이 그 장 글에 남아 있으면 「고칠 것」. 단 실적 · 투입 인력 장은 다른 고객명이 실적 목록일 수 있으니 「확인할 것」으로 둔다.
- `HCG_SRC` · `HCG_DOC` 가 있는 장은 지난 제안서에서 가져온 장이다. 숫자 · 기간 · 인원 · 금액은 지난 고객 값일 수 있으니 「확인할 것」으로 묻는다.
- `HCG_PART` 장은 빈 Component 다. /// 자리가 남았는지 본다. `HCG_PLAN_TYPE` 장은 아직 Library 로 채우지 않은 Dummy 장이다.
- `HCG_REPLACED` 에 있는 바꿈은 이미 끝난 것이다. 다시 바꾸지 않는다.
- 요청 문장의 「N번 장 M번 칸」은 작업창 「칸」 탭의 번호다. 문장에 도형 이름이 같이 오면 도형 이름으로 찾는다.

## Claude for PowerPoint 에서 읽기

`execute_office_js` 로 아래를 돌린다(`context` 는 도구가 넘겨준다).

```js
const p = context.presentation;
p.tags.load("items/key,items/value");
const slides = p.slides;
slides.load("items/id");
await context.sync();
for (const s of slides.items) {
  s.tags.load("items/key,items/value");
  s.shapes.load("items/name");
}
await context.sync();
const client = (p.tags.items.find((t) => t.key === "HCG_CLIENT") || {}).value || "";
const marked = [];
slides.items.forEach((s, i) => {
  const t = {};
  for (const x of s.tags.items) if (x.key.startsWith("HCG_")) t[x.key] = x.value;
  const nums = s.shapes.items.filter((h) => /^HCG_NUM_/.test(h.name)).length;
  if (Object.keys(t).length || nums) marked.push({ no: i + 1, id: s.id, ...t, numBoxes: nums });
});
return { client, marked };
```

## Claude Code(파이썬)에서 읽기

```python
from lxml import etree
from pptx import Presentation

def hcg_tags(part):
    out = {}
    for rel in part.rels.values():
        if rel.reltype.endswith("/tags"):
            for t in etree.fromstring(rel.target_part.blob).iter("{*}tag"):
                if t.get("name", "").startswith("HCG_"):
                    out[t.get("name")] = t.get("val")
    return out

prs = Presentation(path)
client = hcg_tags(prs.part).get("HCG_CLIENT", "")
marked = {i: hcg_tags(s.part) for i, s in enumerate(prs.slides, 1)}
```
