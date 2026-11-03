"""글 바꾸기 + 구조 조작을 plan.json 대로 적용해 새 파일로 저장. 원본은 건드리지 않음.

사용:  python apply.py <pptx> --slide N --plan plan.json --out new.pptx [--log log.json]
  --slide 없이 쓰면 plan 의 「names」(고객명 바꾸기)만 전 장표에 적용함.

plan.json 형식(모든 키 선택):
{
  "names": {"이전 고객명": "새 고객명"},
  "texts": {"3": "새 문장",                      # 칸 번호(slotmap.py 의 no)
            "37": ["문단1", "문단2"],             # 문단별
            "14": [["첫 run", "둘째 run"]],       # 문단 안 run 별(위첨자처럼 서식이 다른 경우)
            "T12.r3c1": "표 12번 셀(3행 1열)",     # 표 칸: T<표 번호>.r<행>c<열> (0부터)
            "id:37": "도형 번호(shape id)로 직접"},
  "ops": [ {"op": "delete_column", "column": 4}, ... ]
}
조작(op) 종류와 인자는 SKILL.md 의 「조작 목록」 참고.
"""
import argparse
import json
import re
import sys

from pptx import Presentation
from pptx.util import Emu

import pptx_ops as O


def die(msg):
    sys.exit('오류: ' + msg)


class Ctx:
    def __init__(self, prs, idx):
        self.prs, self.idx = prs, idx
        self.slide = prs.slides[idx]
        self.SW, self.SH = prs.slide_width, prs.slide_height
        self.px = self.SW / 1600.0
        self.smap, self.live = O.build_slots(prs, idx)
        self.rec = {s['no']: s for s in self.smap['slots']}
        self.log, self.warns = [], []
        self.elmap = {}      # 복제 원본 요소 id → 새 요소
        self.clones = []     # 복제된 새 도형

    def target(self, key):
        """칸 번호·T표.r행c열·id:번호 → 개체(도형 또는 표 칸)"""
        key = str(key)
        m = re.fullmatch(r'T(\d+)\.r(\d+)c(\d+)', key)
        if m:
            no, r, c = map(int, m.groups())
            fr = self.live.get(no)
            if fr is None or not getattr(fr, 'has_table', False):
                die('표를 못 찾음: ' + key)
            return fr.table.cell(r, c)
        if key.startswith('id:'):
            sid = int(key[3:])
            for sh, *_ in O.walk(self.slide.shapes):
                if sh.shape_id == sid:
                    return sh
            die('도형 번호를 못 찾음: ' + key)
        if key.isdigit() and int(key) in self.live:
            return self.live[int(key)]
        die('칸 번호를 못 찾음: %s (slotmap.py 로 만든 번호인지, 장표 번호가 맞는지 확인)' % key)

    def clone_target(self, key):
        """복제된 쪽의 대응 개체(복제 원본 칸 번호로 찾음)"""
        key = str(key)
        m = re.fullmatch(r'T(\d+)\.r(\d+)c(\d+)', key)
        if m:
            no, r, c = map(int, m.groups())
            fr = self.live.get(no)
            nel = self.elmap.get(id(fr._element)) if fr is not None else None
            if nel is None:
                return None
            return O.find_by_element(self.slide, nel).table.cell(r, c)
        obj = self.target(key)
        nel = self.elmap.get(id(obj._element))
        return O.find_by_element(self.slide, nel) if nel is not None else None

    def write(self, obj, paras, label):
        o_n, n_n, o_w, n_w = O.set_paras(obj, paras)
        row = {'칸': label, '원래_글자수': o_n, '새_글자수': n_n,
               '원래_폭가중': round(o_w, 1), '새_폭가중': round(n_w, 1)}
        if o_w:
            r = (n_w - o_w) / o_w * 100
            row['폭증감%'] = round(r)
            if abs(r) > 20 and abs(n_n - o_n) > 2:
                w = '한도 ±20%% 밖: %s (폭가중 %.1f → %.1f, %+d%%) %s' % (
                    label, o_w, n_w, round(r), '넘침 위험 · 줄여 쓰기 권장' if r > 0 else '칸이 비어 보일 수 있음')
                row['경고'] = w
                self.warns.append(w)
        self.log.append(row)


def run_ops(cx, ops):
    for i, op in enumerate(ops, 1):
        name = op.get('op')
        try:
            fn = OPS[name]
        except KeyError:
            die('모르는 조작: %s (op %d)' % (name, i))
        try:
            msg = fn(cx, op)
        except ValueError as e:
            die('조작 %d (%s) 실패: %s' % (i, name, e))
        print('조작 %d %s: %s' % (i, name, msg))


def _apply_texts_clone(cx, texts, what):
    for k, v in (texts or {}).items():
        t = cx.clone_target(k)
        if t is None:
            cx.warns.append('%s: 칸 %s 는 복제된 부분에 없어 글을 못 넣음' % (what, k))
            continue
        cx.write(t, v, '(복제) ' + str(k))


def op_delete_column(cx, op):
    wn, n = O.delete_column(cx.slide, cx.SW, cx.SH, int(op['column']), cx.warns)
    return '열 %d 삭제 → %d열, 열 폭 %d px' % (op['column'], n, round(wn / cx.px))


def op_duplicate_column(cx, op):
    wn, n, em = O.duplicate_column(cx.slide, cx.SW, cx.SH, int(op['column']), cx.warns)
    cx.elmap.update(em)
    _apply_texts_clone(cx, op.get('texts'), '열 복제')
    return '열 %d 복제 → %d열, 열 폭 %d px' % (op['column'], n, round(wn / cx.px))


def op_delete_table_row(cx, op):
    fr = cx.target(op['slot'])
    g = O.detect_gantt(cx.slide, cx.SW, cx.SH)
    g = g if g and g['frame']._element is fr._element else None
    n, each, rb = O.delete_table_rows(cx.slide, fr, op['rows'], int(op.get('header_rows', 1)), g)
    return '표 행 %s 삭제 → %d행, 행 높이 %d px%s' % (op['rows'], n, round(each / cx.px),
                                                   (' · 막대 %d개 함께 삭제' % rb) if g else '')


def op_gantt_bar(cx, op):
    g = O.detect_gantt(cx.slide, cx.SW, cx.SH)
    if not g:
        raise ValueError('간트 표(주차 열 4개 이상)를 못 찾음')
    sh = cx.target(op['slot'])
    s, e = int(op['start']), int(op['end'])
    if not (1 <= s <= e <= g['weeks']):
        raise ValueError('주차는 1~%d, 시작 ≤ 끝이어야 함' % g['weeks'])
    if op.get('copy'):
        sh, em = O.clone(cx.slide, sh, name=op.get('name'))
        cx.clones.append(sh)
    O.set_bar(sh, g['x0'], g['wk'], s, e)
    return '막대 %s → %d~%d주%s' % (op['slot'], s, e, ' (복제)' if op.get('copy') else '')


def op_duplicate_box(cx, op):
    slots = op['slots']
    dx = float(op.get('dx_px', 0)) * cx.px
    dy = float(op.get('dy_px', 0)) * cx.px
    if op.get('to_week'):
        g = O.detect_gantt(cx.slide, cx.SW, cx.SH)
        if not g:
            raise ValueError('to_week 는 간트 표가 있어야 함')
        first = cx.target(slots[0])
        lk = O.shape_lookup(cx.slide)[id(first._element)]
        cxm = lk[1][0] + lk[1][2] / 2
        dx = g['x0'] + int(op['to_week']) * g['wk'] - cxm
    new_all = []
    for k in slots:
        sh = cx.target(k)
        ns, em = O.clone(cx.slide, sh, dx=dx, dy=dy, name=op.get('name') if k == slots[0] else None)
        cx.elmap.update(em)
        cx.clones.append(ns)
        new_all.append(ns)
    _apply_texts_clone(cx, op.get('texts'), '박스 복제')
    first = cx.clone_target(slots[0])
    for key, attr in (('left_px', 'left'), ('top_px', 'top'), ('w_px', 'width'), ('h_px', 'height')):
        if key in op and first is not None:
            setattr(first, attr, int(op[key] * cx.px))
    for k in slots:
        ns = cx.clone_target(k)
        if ns is None:
            continue
        if op.get('fill') and hasattr(ns, 'fill'):
            O.set_fill(ns, op['fill'])
        if op.get('text_color') and ns.has_text_frame:
            O.set_text_color(ns, op['text_color'])
    if op.get('overlay'):
        cx.clones = [c for c in cx.clones if c not in new_all]
    return '%d개 복제(이름에 「(추가)」), 이동 %+d,%+d px' % (len(slots), round(dx / cx.px), round(dy / cx.px))


def op_move(cx, op):
    sh = cx.target(op['slot'])
    if 'dx_px' in op:
        sh.left = sh.left + int(op['dx_px'] * cx.px)
    if 'dy_px' in op:
        sh.top = sh.top + int(op['dy_px'] * cx.px)
    if 'w_px' in op:
        sh.width = int(op['w_px'] * cx.px)
    if 'h_px' in op:
        sh.height = int(op['h_px'] * cx.px)
    return '칸 %s 이동·크기 변경' % op['slot']


def op_fill(cx, op):
    sh = cx.target(op['slot'])
    if op.get('color'):
        O.set_fill(sh, op['color'])
    if op.get('text_color'):
        O.set_text_color(sh, op['text_color'])
    return '칸 %s 색 변경' % op['slot']


def op_shape(cx, op):
    sh = cx.target(op['slot'])
    xf = sh._element.spPr.find(O.qa('xfrm'))
    if xf is not None and 'rot' in xf.attrib and op.get('clear_rotation', True):
        del xf.attrib['rot']
    if op.get('prst'):
        sh._element.spPr.find(O.qa('prstGeom')).set('prst', op['prst'])
    for k_, a in (('left_px', 'left'), ('top_px', 'top'), ('w_px', 'width'), ('h_px', 'height')):
        if k_ in op:
            setattr(sh, a, int(op[k_] * cx.px))
    if op.get('color'):
        O.set_fill(sh, op['color'])
    return '칸 %s 모양 변경' % op['slot']


def op_line(cx, op):
    sh = cx.target(op['slot'])
    b, e = op['begin_px'], op['end_px']
    sh.begin_x, sh.begin_y = int(b[0] * cx.px), int(b[1] * cx.px)
    sh.end_x, sh.end_y = int(e[0] * cx.px), int(e[1] * cx.px)
    return '선 %s 끝점 이동' % op['slot']


def op_delete_shape(cx, op):
    O.delete_shape(cx.target(op['slot']))
    return '칸 %s 삭제' % op['slot']


def op_font_size(cx, op):
    d = int(op['delta_pt'])
    if not op.get('approved'):
        raise ValueError('글꼴 줄이기는 사용자 승인(approved: true) 뒤에만 함')
    if d > 0 or d < -2:
        raise ValueError('글꼴은 1pt 씩, 최대 -2pt 까지만 줄임')
    sh = cx.target(op['slot'])
    n = 0
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                r.font.size = Emu(r.font.size + d * 12700)
                n += 1
    if not n:
        cx.warns.append('칸 %s: 글꼴 크기가 상속이라 못 줄임(run 에 크기 없음)' % op['slot'])
    return '칸 %s 글꼴 %dpt (run %d개)' % (op['slot'], d, n)


OPS = {'delete_column': op_delete_column, 'duplicate_column': op_duplicate_column,
       'delete_table_row': op_delete_table_row, 'gantt_bar': op_gantt_bar,
       'duplicate_box': op_duplicate_box, 'move': op_move, 'fill': op_fill, 'shape': op_shape,
       'line': op_line, 'delete_shape': op_delete_shape, 'font_size': op_font_size}


def overlap_warnings(cx):
    """복제한 도형의 글상자가 다른 글상자와 25% 넘게 겹치면 경고"""
    if not cx.clones:
        return
    look = O.shape_lookup(cx.slide)
    texts = [(v[0], v[1]) for v in look.values() if v[0].has_text_frame and v[0].text_frame.text.strip()
             and not O.is_group(v[0])]
    cl_els = set()
    for c in cx.clones:
        for el in c._element.iter():
            cl_els.add(id(el))
    for a, ba in texts:
        if id(a._element) not in cl_els:
            continue
        for b, bb in texts:
            if b is a or id(b._element) in cl_els:
                continue
            ix = min(ba[0] + ba[2], bb[0] + bb[2]) - max(ba[0], bb[0])
            iy = min(ba[1] + ba[3], bb[1] + bb[3]) - max(ba[1], bb[1])
            if ix > 0 and iy > 0 and ix * iy > 0.25 * min(ba[2] * ba[3], bb[2] * bb[3]):
                cx.warns.append('겹침: 복제한 「%s」 가 「%s」 와 겹침 · 글을 줄이거나 위치를 옮길 것' % (a.name, b.name))


def main():
    ap = argparse.ArgumentParser(description='plan.json 대로 글·구조 바꾸기(새 파일로 저장)')
    ap.add_argument('pptx')
    ap.add_argument('--slide', type=int, help='고칠 장표 번호(1부터)')
    ap.add_argument('--plan', required=True)
    ap.add_argument('--out', required=True, help='새 파일 경로(원본과 달라야 함)')
    ap.add_argument('--log', help='글자 수 대조 기록을 JSON 으로 저장')
    a = ap.parse_args()
    if a.out.replace('\\', '/').lower() == a.pptx.replace('\\', '/').lower():
        die('원본을 덮어쓰지 않음 · --out 은 새 이름으로')
    plan = json.load(open(a.plan, encoding='utf-8'))
    prs = Presentation(a.pptx)
    names = plan.get('names') or {}
    if not a.slide:
        if plan.get('texts') or plan.get('ops'):
            die('texts·ops 는 --slide 가 있어야 함')
        n = 0
        for i in range(len(prs.slides)):
            n += O.replace_names(prs.slides[i], names)
        n += O.replace_props(prs, names)
        prs.save(a.out)
        print('고객명 %d곳 바꿈 → %s' % (n, a.out))
        return
    if not 1 <= a.slide <= len(prs.slides):
        die('장표 번호 범위 밖: %d (전체 %d장)' % (a.slide, len(prs.slides)))
    cx = Ctx(prs, a.slide - 1)
    # 대상 먼저 확인(오류가 있으면 아무것도 바꾸지 않고 끝냄)
    for k in (plan.get('texts') or {}):
        cx.target(k)
    if names:
        n = O.replace_names(cx.slide, names) + O.replace_props(prs, names)
        print('고객명 %d곳 바꿈' % n)
    for k, v in (plan.get('texts') or {}).items():
        rec = cx.rec.get(int(k)) if str(k).isdigit() else None
        cx.write(cx.target(k), v, '%s %s' % (k, (rec or {}).get('name', '')))
    run_ops(cx, plan.get('ops') or [])
    overlap_warnings(cx)
    prs.save(a.out)
    print('글 바꾼 칸 %d개 · 경고 %d건 → %s' % (len(cx.log), len(cx.warns), a.out))
    for w in cx.warns:
        print('  경고:', w)
    if a.log:
        json.dump({'log': cx.log, 'warnings': cx.warns}, open(a.log, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
