"""칸 지도 만들기: 장표마다 칸 번호 · 이름 · 종류 · 역할 추정 · 위치 · 글자 수 · 글꼴 크기 · 폭 가중 글자 수 한도.

사용:  python slotmap.py <pptx> [--slide N] --out slots.json
  --slide N  : N번 장표만(1부터). 없으면 전체 장표.
결과 slots.json 은 {"slides":[{slide, layout, slots:[...], columns, gantt}]} 형태.
역할은 위치·글꼴 크기·길이로 짐작한 값이므로 사람(또는 Claude)이 한 번 확인할 것.
"""
import argparse
import json
import sys

from pptx import Presentation

from pptx_ops import build_slots


def main():
    ap = argparse.ArgumentParser(description='칸 지도(slots.json) 만들기')
    ap.add_argument('pptx')
    ap.add_argument('--slide', type=int, help='장표 번호(1부터). 없으면 전체')
    ap.add_argument('--out', required=True, help='저장할 slots.json 경로')
    a = ap.parse_args()
    prs = Presentation(a.pptx)
    n = len(prs.slides)
    idxs = [a.slide - 1] if a.slide else range(n)
    if a.slide and not 1 <= a.slide <= n:
        sys.exit('장표 번호 범위 밖: %d (전체 %d장)' % (a.slide, n))
    res = []
    for i in idxs:
        m, _ = build_slots(prs, i)
        res.append(m)
        extra = []
        if 'columns' in m:
            extra.append('열 %d개' % len(m['columns']))
        if 'gantt' in m:
            extra.append('간트 %d주 · 막대 %d개' % (m['gantt']['weeks'], len(m['gantt']['bars'])))
        print('장표 %d  칸 %d개  %s' % (i + 1, len(m['slots']), ' · '.join(extra)))
    with open(a.out, 'w', encoding='utf-8') as f:
        json.dump({'source': a.pptx, 'slides': res}, f, ensure_ascii=False, indent=1)
    print('저장:', a.out)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
