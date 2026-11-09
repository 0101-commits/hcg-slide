# Dummy 제안서 만들기

사용자가 「Dummy 제안서」「더미 만들어 줘」라고 하며 고객사명 · 회의 내용 · 주요 스콥을 주면, 목차와 장마다 유형 · 어깨끈 · 핵심 문장을 정해 **Dummy 장**을 만든다. 실제 장표는 「HCG 장표」 작업창의 「Library로 채우기」가 유형마다 지난 제안서 장표로 바꿔 넣는다.

## 1. 입력 확인
- 꼭 필요: 고객사명, 업종(금융 · 자동차·기계·중공업 · 식품·소비재 · 지주·그룹 · IT·플랫폼·게임 · 화학·소재·에너지 · 공공·연구 · 건설·부동산 · 바이오·제약 · 유통·서비스·레저 · 기타 중), 회의 내용(메모), 주요 스콥, 기간, 분량(일반 25장 · 약식 12장).
- 빠진 것만 묻는다. 분량을 말하지 않으면 일반 25장.

## 2. 목차 · 장 계획
- 목차는 `toc_and_messages.md` 의 유형 중 하나, 장 흐름은 Ⅰ 제안 개요 → Ⅱ 수행 방안 → Ⅲ 수행 체계 → Ⅳ HCG 소개.
- 장마다 `types.md` 의 유형 코드 하나(T101~T604)를 고른다. 같은 유형이 여러 장이어도 됨(예: 가설적 이슈 3장).
- 어깨끈 = 「대분류    소분류」(예: 「제안 개요    가설적 이슈」), 핵심 문장 = 두 줄 안 결론형 「~함」. 회의 내용에서 나온 이슈 · 스콥을 핵심 문장에 담는다. 지어낸 숫자 금지(필요하면 「(가정)」).
- 제도 설계 장(T211~T216)은 스콥에 맞는 HR 모듈을 mod 에 적는다(평가 · 보상 · 직급 · 직무 · 조직 · 육성 · 조직문화 · 인력 · 시스템 · 채용 · 노사 · 공통 중).
- 가격 장표는 만들지 않는다.

## 3-A. Claude for PowerPoint 에서 — Dummy 장을 직접 만든다
Office 명령 실행 도구(`execute_office_js`)로 맨 뒤에 장을 하나씩 더한다. 장마다 아래 표시(태그)를 꼭 붙인다 — 작업창이 이 표시로 Library 장표를 고른다.

```js
await PowerPoint.run(async (ctx) => {
  ctx.presentation.tags.add("HCG_CLIENT", "한빛소재");
  const plan = [ /* {ch, head, msg, type, mod} … */ ];
  for (const s of plan) {
    ctx.presentation.slides.add();
    await ctx.sync();
    const all = ctx.presentation.slides; all.load("items"); await ctx.sync();
    const sl = all.items[all.items.length - 1];
    const t1 = sl.shapes.addTextBox(`${s.ch}    ${s.head}`, { left: 36, top: 18, width: 880, height: 26 });
    t1.textFrame.textRange.font.size = 14; t1.textFrame.textRange.font.color = "#C00000"; t1.textFrame.textRange.font.bold = true;
    const t2 = sl.shapes.addTextBox(s.msg, { left: 36, top: 46, width: 880, height: 52 });
    t2.textFrame.textRange.font.size = 18; t2.textFrame.textRange.font.bold = true;
    sl.shapes.addTextBox("///", { left: 36, top: 120, width: 880, height: 300 });
    sl.tags.add("HCG_PLAN_TYPE", s.type);
    sl.tags.add("HCG_PLAN_HEAD", s.head);
    sl.tags.add("HCG_PLAN_MSG", s.msg.slice(0, 250));
    if (s.mod) sl.tags.add("HCG_PLAN_MOD", s.mod);
    await ctx.sync();
  }
});
```
- 한 번에 10장씩 나눠 실행해도 된다. 끝나면 「HCG 장표 → Library → Dummy 제안서 › → 모두 Library로 채우기」를 누르라고 안내한다.
- 도구가 없거나 실패하면 3-B 로.

## 3-B. Claude 앱 · Claude Code 에서(또는 3-A 가 안 될 때) — 계획 블록을 준다
답 끝에 아래 블록을 그대로 준다. 사용자가 HCG 장표 작업창 「Dummy 제안서 › 계획 붙여넣기」에 붙여넣으면 작업창이 Dummy 장을 만든다.

```
HCG_PLAN {"client": "한빛소재", "industry": "화학·소재·에너지", "slides": [
 {"ch": "제안 개요", "head": "제안 배경", "type": "T101", "msg": "…함"},
 {"ch": "제안 개요", "head": "가설적 이슈", "type": "T105", "msg": "…함", "mod": "보상"}
]}
```
- JSON 은 한 덩어리로, 큰따옴표만. 장 수만큼 slides 에 넣는다.

## 4. 그 뒤
- 채운 뒤 사용자가 「Claude에게 ›」로 넘기면 `hcg-restructure` 로 회의 메모 기준 글 고치기, 마지막에 `hcg-check`.
