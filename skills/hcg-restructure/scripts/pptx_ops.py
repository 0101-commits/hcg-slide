"""hcg-restructure 공통 부품: 칸 지도 · 글자 교체 · 도형 조작 (python-pptx + lxml).

slotmap.py · apply.py 가 가져다 씀. 단독 실행 안 함.
좌표 단위 「px」 = 장표 가로를 1600 으로 본 좌표(slots.json 의 bbox_px 와 같음).
"""
import copy
import re

from lxml import etree
from pptx.util import Emu

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'


def qa(tag):
    return '{%s}%s' % (NS_A, tag)


def qp(tag):
    return '{%s}%s' % (NS_P, tag)


def local(el):
    return etree.QName(el).localname


# ---------------------------------------------------------------- 글자 수

def nsp(s):
    """공백 뺀 글자 수"""
    return len(re.sub(r'\s', '', s))


def wlen(s):
    """폭 가중 글자 수: 한글·한자 1 · 대문자 0.7 · 소문자·숫자 0.55 · 문장부호 0.35 · 공백 0.3"""
    w = 0.0
    for ch in s:
        if ch.isspace():
            w += 0.3
        elif '가' <= ch <= '힣' or '一' <= ch <= '鿿' or ch in '○◯':
            w += 1
        elif ch.isupper():
            w += 0.7
        elif ch.isalnum():
            w += 0.55
        else:
            w += 0.35
    return w


# ---------------------------------------------------------------- 도형 훑기

def is_group(sh):
    return sh._element.tag == qp('grpSp')


def walk(shapes, T=(1.0, 0.0, 1.0, 0.0), depth=0, path=()):
    """(도형, 절대 bbox EMU(l,t,w,h), 깊이, 묶음 경로, 부모 변환 T) 를 차례로 내줌.
    T=(sx,ox,sy,oy): 부모 좌표 → 절대 좌표."""
    for sh in shapes:
        sx, ox, sy, oy = T
        try:
            l = ox + sh.left * sx
            t = oy + sh.top * sy
            w = sh.width * sx
            h = sh.height * sy
        except Exception:
            l = t = w = h = 0.0
        yield sh, (l, t, w, h), depth, path, T
        if is_group(sh):
            x = sh._element.grpSpPr.xfrm
            if x is None:
                continue
            co, ce = x.chOff, x.chExt
            nsx = (w / ce.cx) if ce.cx else 1.0
            nsy = (h / ce.cy) if ce.cy else 1.0
            nT = (nsx, l - co.x * nsx, nsy, t - co.y * nsy)
            yield from walk(sh.shapes, nT, depth + 1, path + (sh.shape_id,))


def kind(sh):
    tag = local(sh._element)
    if tag == 'grpSp':
        return 'group'
    if tag == 'graphicFrame':
        if getattr(sh, 'has_table', False) and sh.has_table:
            return 'table'
        if getattr(sh, 'has_chart', False) and sh.has_chart:
            return 'chart'
        return 'object'
    if tag == 'pic':
        return 'picture'
    if tag == 'cxnSp':
        return 'line'
    if sh.has_text_frame and sh.text_frame.text.strip():
        return 'text'
    return 'shape'


def prst_of(sh):
    try:
        g = sh._element.spPr.find(qa('prstGeom'))
        return g.get('prst') if g is not None else None
    except Exception:
        return None


def font_pt(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            if r.font.size:
                return round(r.font.size.pt, 1)
    return None


def shape_lookup(slide):
    """요소 → (도형, bbox, 깊이, 경로, T)"""
    return {id(sh._element): (sh, bb, dep, path, T)
            for sh, bb, dep, path, T in walk(slide.shapes)}


def find_by_element(slide, el):
    for sh, bb, dep, path, T in walk(slide.shapes):
        if sh._element is el:
            return sh
    raise KeyError('요소를 못 찾음')


def next_id(slide):
    mx = 0
    for sh, *_ in walk(slide.shapes):
        mx = max(mx, sh.shape_id)
    return mx + 1


# ---------------------------------------------------------------- 칸 지도

def guess_role(rec, W, H):
    """위치·글꼴 크기·길이로 짐작한 역할(사람이 한 번 확인할 것)"""
    k = rec['kind']
    l, t, w, h = rec['bbox_px']
    if k == 'table':
        return '표'
    if k == 'cell':
        if rec['r'] == 0:
            return '표 머리 행'
        return '표 열 머리' if rec['c'] == 0 else '표 본문'
    if k == 'picture':
        return '아이콘·로고 후보' if (w < 0.1 * W and h < 0.1 * W) else '그림'
    if k == 'line':
        return '선·연결선'
    if k == 'group':
        return '묶음'
    if k == 'chart':
        return '차트'
    if k == 'object':
        return '개체(SmartArt 등)'
    if k == 'shape':
        prst = rec.get('prst') or ''
        if h <= 0.04 * H and w >= 0.03 * W:
            return '막대 후보'
        if 'rrow' in prst or 'chevron' in prst or 'triangle' in prst or 'homePlate' in prst:
            return '방향 도형'
        return '도형'
    # text
    fs = rec.get('font_pt') or 0
    wc = rec.get('wchars', 0)
    txt = rec.get('text', '')
    if t < 0.1 * H and wc <= 22:
        return '분류 꼬리표·제목'
    if t < 0.26 * H and wc >= 15 and (fs >= 16 or fs == 0):
        return '핵심 문장'
    if t + h > 0.9 * H and (fs and fs <= 10 or txt.lstrip().startswith(('※', '*', '출처', '1)'))):
        return '각주'
    if fs >= 16 and wc <= 25:
        return '큰 글자·소제목'
    if wc <= 22:
        return '머리·소제목'
    return '본문'


def _px(v, k):
    return round(v * k)


def build_slots(prs, idx):
    """장표 하나의 칸 지도. (지도 dict, {번호: 개체}) 를 돌려줌. idx 는 0부터."""
    slide = prs.slides[idx]
    SW, SH = prs.slide_width, prs.slide_height
    k = 1600.0 / SW
    W, H = 1600, SH * k
    items = []
    for sh, bb, dep, path, T in walk(slide.shapes):
        if sh.is_placeholder:
            try:
                pt = str(sh.placeholder_format.type)
            except Exception:
                pt = ''
            if any(x in pt for x in ('SLIDE_NUMBER', 'FOOTER', 'DATE')):
                continue
        kd = kind(sh)
        if sh.is_placeholder and kd == 'shape':
            continue
        l, t, w, h = bb
        rec = {'id': sh.shape_id, 'name': sh.name, 'kind': kd,
               'group_path': '/'.join(map(str, path)) or None,
               'bbox_px': [_px(l, k), _px(t, k), _px(w, k), _px(h, k)]}
        if kd == 'shape':
            rec['prst'] = prst_of(sh)
        if sh.has_text_frame and kd in ('text', 'shape') and sh.text_frame.text.strip():
            tx = sh.text_frame.text.replace('\x0b', '\n')
            rec.update(chars=nsp(tx), wchars=round(wlen(tx), 1), font_pt=font_pt(sh.text_frame),
                       text=tx[:120], paragraphs=len(sh.text_frame.paragraphs))
            rec['limit'] = {'min': round(rec['wchars'] * 0.8, 1), 'max': round(rec['wchars'] * 1.2, 1)}
        if kd == 'picture':
            try:
                rec['descr'] = sh._element.nvPicPr.cNvPr.get('descr')
            except Exception:
                pass
        if kd == 'table':
            tb = sh.table
            merged = any(c.is_merge_origin or c.is_spanned for r in tb.rows for c in r.cells)
            rec['table'] = {'rows': len(tb.rows), 'cols': len(tb.columns), 'merged': merged}
        rec['role_guess'] = None
        items.append([sh, rec, bb, None])
        if kd == 'table':
            tb = sh.table
            xs = [l]
            for col in tb.columns:
                xs.append(xs[-1] + col.width * T[0])
            ys = [t]
            for row in tb.rows:
                ys.append(ys[-1] + row.height * T[2])
            for r in range(len(tb.rows)):
                for c in range(len(tb.columns)):
                    cell = tb.cell(r, c)
                    if cell.is_spanned:
                        continue
                    tx = cell.text_frame.text.replace('\x0b', '\n')
                    if not tx.strip():
                        continue
                    x2 = xs[min(c + cell.span_width, len(tb.columns))]
                    y2 = ys[min(r + cell.span_height, len(tb.rows))]
                    crec = {'id': sh.shape_id, 'name': sh.name, 'kind': 'cell',
                            'group_path': '/'.join(map(str, path)) or None,
                            'bbox_px': [_px(xs[c], k), _px(ys[r], k), _px(x2 - xs[c], k), _px(y2 - ys[r], k)],
                            'r': r, 'c': c, 'rowspan': cell.span_height, 'colspan': cell.span_width,
                            'chars': nsp(tx), 'wchars': round(wlen(tx), 1),
                            'font_pt': font_pt(cell.text_frame), 'text': tx[:120],
                            'paragraphs': len(cell.text_frame.paragraphs), 'role_guess': None}
                    crec['limit'] = {'min': round(crec['wchars'] * 0.8, 1), 'max': round(crec['wchars'] * 1.2, 1)}
                    items.append([cell, crec, (xs[c], ys[r], x2 - xs[c], ys[r + cell.span_height] - ys[r]), sh])
    items.sort(key=lambda it: (round(it[1]['bbox_px'][1] / 45), it[1]['bbox_px'][0]))
    live = {}
    no_of_el = {}
    for n, (obj, rec, bb, fr) in enumerate(items, 1):
        rec['no'] = n
        rec['role_guess'] = guess_role(rec, W, H)
        live[n] = obj
        if rec['kind'] != 'cell':
            no_of_el[id(obj._element)] = n
    for obj, rec, bb, fr in items:
        if rec['kind'] == 'cell':
            rec['table_no'] = no_of_el.get(id(fr._element))
        if rec['kind'] == 'group':
            rec['children'] = [no_of_el.get(id(c._element)) for c in obj.shapes
                               if id(c._element) in no_of_el]
    # no 를 맨 앞으로
    slots = []
    for obj, rec, bb, fr in items:
        slots.append({'no': rec['no'], **{kk: vv for kk, vv in rec.items() if kk != 'no'}})
    out = {'slide': idx + 1, 'layout': slide.slide_layout.name,
           'size_emu': [SW, SH], 'px_note': '좌표 px = 장표 가로를 1600 으로 본 값',
           'hidden': slide._element.get('show') == '0', 'slots': slots}
    cols = detect_columns(slide, SW, SH)
    if cols:
        out['columns'] = [{'column': i + 1, 'x_px': [_px(c['l'], k), _px(c['r'], k)],
                           'slots': sorted(no_of_el[id(u[0]._element)] for u in c['units']
                                           if id(u[0]._element) in no_of_el)}
                          for i, c in enumerate(cols['cols'])]
        out['column_separators'] = [no_of_el.get(id(s[0]._element)) for s in cols['seps']]
    g = detect_gantt(slide, SW, SH)
    if g:
        out['gantt'] = {'table_no': no_of_el.get(id(g['frame']._element)), 'weeks': g['weeks'],
                        'week_px': _px(g['wk'], k), 'x0_px': _px(g['x0'], k),
                        'bars': [{'slot': no_of_el.get(id(b[0]._element)), 'start': b[2], 'end': b[3], 'row': b[4]}
                                 for b in g['bars']]}
    return out, live


# ---------------------------------------------------------------- 글자 교체

def _mk_run(proto_r, text):
    r = copy.deepcopy(proto_r)
    r.find(qa('t')).text = text
    rpr = r.find(qa('rPr'))
    if rpr is None:
        rpr = etree.Element(qa('rPr'))
        r.insert(0, rpr)
    rpr.set('lang', 'ko-KR')
    if 'altLang' in rpr.attrib:
        del rpr.attrib['altLang']
    return r


def _proto_from_empty(pe):
    r = etree.Element(qa('r'))
    rpr = etree.SubElement(r, qa('rPr'))
    end = pe.find(qa('endParaRPr'))
    if end is not None:
        rpr.attrib.update(end.attrib)
        for ch in end:
            rpr.append(copy.deepcopy(ch))
    etree.SubElement(r, qa('t')).text = ''
    return r


def set_para(p, text):
    """문단 하나의 글자 교체. 문단 서식·첫 run 서식은 그대로, 새 run 에 lang=ko-KR.
    text 가 str 이면 첫 run 서식으로 전체, list 면 run 별(서식 순서대로). '\\n' 은 줄바꿈."""
    pe = p._p
    runs = [c for c in pe if c.tag == qa('r')]
    brproto = pe.find(qa('br'))
    if runs:
        pos = list(pe).index(runs[0])
    else:
        end = pe.find(qa('endParaRPr'))
        pos = list(pe).index(end) if end is not None else len(pe)
        runs = [_proto_from_empty(pe)]
    for c in list(pe):
        if c.tag in (qa('r'), qa('br')):
            pe.remove(c)
    is_list = isinstance(text, list)
    for i, part in enumerate(text if is_list else [text]):
        proto = runs[min(i, len(runs) - 1)] if is_list else runs[0]
        if part == '' and is_list:
            continue
        for j, ln in enumerate(part.split('\n')):
            if j > 0:
                br = copy.deepcopy(brproto) if brproto is not None else etree.Element(qa('br'))
                if brproto is None and proto.find(qa('rPr')) is not None:
                    br.append(copy.deepcopy(proto.find(qa('rPr'))))
                pe.insert(pos, br)
                pos += 1
            if ln == '' and j > 0:
                continue
            pe.insert(pos, _mk_run(proto, ln))
            pos += 1


def set_paras(target, paras):
    """칸(도형·표 칸) 하나의 글 교체. paras: 문단별 str|list. 문단 수가 다르면 마지막 문단을 복제/삭제.
    (원래 글자 수, 새 글자 수, 원래 폭가중, 새 폭가중) 를 돌려줌."""
    tf = target.text_frame
    o_n, o_w = nsp(tf.text), wlen(tf.text)
    if isinstance(paras, str):
        paras = [paras]
    ps = list(tf.paragraphs)
    txbody = ps[0]._p.getparent()
    while len(ps) < len(paras):
        newp = copy.deepcopy(ps[-1]._p)
        ps[-1]._p.addnext(newp)
        ps = list(tf.paragraphs)
    while len(ps) > len(paras):
        txbody.remove(ps[-1]._p)
        ps = list(tf.paragraphs)
    for p, t in zip(ps, paras):
        set_para(p, t)
    return o_n, nsp(tf.text), o_w, wlen(tf.text)


# ---------------------------------------------------------------- 이름 바꾸기

def _replace_in_para(pe, mapping):
    n = 0
    runs = [r for r in pe if r.tag == qa('r')]
    if not runs:
        return 0
    for old in sorted(mapping, key=len, reverse=True):
        new = mapping[old]
        texts = [r.find(qa('t')).text or '' for r in runs]
        full = ''.join(texts)
        bounds, pos = [], 0
        for t in texts:
            bounds.append((pos, pos + len(t)))
            pos += len(t)
        for m in reversed(list(re.finditer(re.escape(old), full))):
            s, e = m.span()
            a = next(i for i, (b0, b1) in enumerate(bounds) if b0 <= s < b1)
            b = next(i for i, (b0, b1) in enumerate(bounds) if b0 < e <= b1)
            if a == b:
                texts[a] = texts[a][:s - bounds[a][0]] + new + texts[a][e - bounds[a][0]:]
            else:
                texts[a] = texts[a][:s - bounds[a][0]] + new
                for i in range(a + 1, b):
                    texts[i] = ''
                texts[b] = texts[b][e - bounds[b][0]:]
            n += 1
            # 경계를 다시 계산하지 않기 위해 한 번에 한 이름씩 처리하고 끝냄
            full = ''.join(texts)
            bounds, pos = [], 0
            for t in texts:
                bounds.append((pos, pos + len(t)))
                pos += len(t)
        for r, t in zip(runs, texts):
            r.find(qa('t')).text = t
    return n


def replace_names(slide, mapping, with_notes=True):
    """장표의 글·표·묶음 이름·그림 설명·발표자 메모에서 이전 고객명 → 새 고객명. 바꾼 횟수를 돌려줌."""
    n = 0
    roots = [slide._element]
    if with_notes and slide.has_notes_slide:
        roots.append(slide.notes_slide._element)
    for root in roots:
        for pe in root.iter(qa('p')):
            n += _replace_in_para(pe, mapping)
        for el in root.iter():
            if isinstance(el.tag, str) and el.tag.endswith('}cNvPr'):
                for at in ('name', 'descr', 'title'):
                    v = el.get(at)
                    if v:
                        nv = v
                        for old in sorted(mapping, key=len, reverse=True):
                            nv = nv.replace(old, mapping[old])
                        if nv != v:
                            el.set(at, nv)
                            n += 1
    return n


def replace_props(prs, mapping):
    cp = prs.core_properties
    n = 0
    for at in ('title', 'subject', 'keywords', 'comments', 'category', 'author', 'last_modified_by'):
        v = getattr(cp, at, None)
        if v:
            nv = v
            for old in sorted(mapping, key=len, reverse=True):
                nv = nv.replace(old, mapping[old])
            if nv != v:
                setattr(cp, at, nv)
                n += 1
    return n


# ---------------------------------------------------------------- 복제·이동·서식

def clone(slide, shape, dx=0, dy=0, name=None, suffix=' (추가)'):
    """도형(묶음 포함) 복제. 새 번호를 주고 같은 부모 안 원본 바로 뒤에 붙임.
    (새 도형, {원본 요소 id: 새 요소}) 를 돌려줌. dx, dy 는 EMU."""
    el = copy.deepcopy(shape._element)
    mx = next_id(slide)
    for i, c in enumerate(el.iter(qp('cNvPr'))):
        c.set('id', str(mx + i))
        if i == 0:
            c.set('name', name or (c.get('name', '') + suffix))
    shape._element.addnext(el)
    elmap = {id(a): b for a, b in zip(shape._element.iter(), el.iter())}
    new = find_by_element(slide, el)
    if dx:
        new.left = new.left + Emu(int(dx))
    if dy:
        new.top = new.top + Emu(int(dy))
    return new, elmap


def delete_shape(sh):
    sh._element.getparent().remove(sh._element)


def set_fill(shape, hexrgb):
    from pptx.dml.color import RGBColor
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(hexrgb)


def set_text_color(shape, hexrgb):
    from pptx.dml.color import RGBColor
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.color.rgb = RGBColor.from_string(hexrgb)


def set_abs(sh, T, l=None, w=None):
    sx, ox, _, _ = T
    if l is not None:
        sh.left = int(round((l - ox) / sx))
    if w is not None:
        sh.width = int(round(w / sx))


def fit_group_x(g, parentT):
    """묶음의 가로 틀을 자식 bbox 에 다시 맞춤(자식 좌표는 그대로, 묶음 변환만 갱신)"""
    xf = g._element.grpSpPr.xfrm
    kids = [k for k in g.shapes]
    if xf is None or not kids:
        return
    cl = min(k.left for k in kids)
    cr = max(k.left + k.width for k in kids)
    co, ce = xf.chOff, xf.chExt
    sx_g = (g.width / ce.cx) if ce.cx else 1.0
    # 자식 좌표 → 부모 좌표
    new_left = g.left + (cl - co.x) * sx_g
    new_w = (cr - cl) * sx_g
    xf.off.x = int(round(new_left))
    xf.ext.cx = int(round(new_w))
    xf.chOff.x = int(cl)
    xf.chExt.cx = int(cr - cl)


# ---------------------------------------------------------------- 열 찾기

def detect_columns(slide, SW, SH):
    """같은 가로 범위의 요소를 한 열로 묶음. 열이 2개 미만이면 None.
    돌려주는 값: {'cols': [{'l','r','units':[(도형,bbox,깊이,경로,T)]}], 'seps': [...], 'span': (L,R)}"""
    W, H = SW, SH
    units, seps = [], []

    def visit(shapes, T, depth, path):
        for sh in shapes:
            sx, ox, sy, oy = T
            try:
                l = ox + sh.left * sx
                t = oy + sh.top * sy
                w = sh.width * sx
                h = sh.height * sy
            except Exception:
                continue
            if t < 0.18 * H or t > 0.93 * H:
                continue
            if is_group(sh):
                if w >= 0.12 * W:
                    x = sh._element.grpSpPr.xfrm
                    co, ce = x.chOff, x.chExt
                    nsx = (w / ce.cx) if ce.cx else 1.0
                    nsy = (h / ce.cy) if ce.cy else 1.0
                    visit(sh.shapes, (nsx, l - co.x * nsx, nsy, t - co.y * nsy), depth + 1, path + (sh.shape_id,))
                else:
                    units.append((sh, (l, t, w, h), depth, path, T))
                continue
            if w < 0.008 * W and h > 0.08 * H:
                seps.append((sh, (l, t, w, h), depth, path, T))
                continue
            if w > 0.6 * W or (w < 0.01 * W and h < 0.01 * H):
                continue
            units.append((sh, (l, t, w, h), depth, path, T))

    visit(slide.shapes, (1.0, 0.0, 1.0, 0.0), 0, ())
    units.sort(key=lambda u: u[1][0])
    clusters = []
    for u in units:
        l, t, w, h = u[1]
        r = l + w
        best, bov = None, 0
        for c in clusters:
            ov = min(r, c['r']) - max(l, c['l'])
            if ov > bov:
                best, bov = c, ov
        if best is not None and bov >= 0.5 * min(w, best['r'] - best['l']):
            best['units'].append(u)
            best['l'], best['r'] = min(best['l'], l), max(best['r'], r)
        else:
            clusters.append({'l': l, 'r': r, 'units': [u]})
    clusters = [c for c in clusters if len(c['units']) >= 3]
    if len(clusters) < 2:
        return None
    clusters.sort(key=lambda c: c['l'])
    seps.sort(key=lambda s: s[1][0])
    return {'cols': clusters, 'seps': seps, 'span': (clusters[0]['l'], clusters[-1]['r'])}


def _relayout(slide, order, det, warns, dirty=()):
    """order: [(원래 l, 원래 폭, [unit...])] 새 열 순서. 전체 폭을 N등분해 다시 놓음."""
    L, R = det['span']
    N = len(order)
    Wn = (R - L) / N
    touched = {}
    parents = list(dirty)
    for i, (ol, ow, us) in enumerate(order):
        nl = L + i * Wn
        for sh, bb, dep, path, T in us:
            l, t, w, h = bb
            rw = w / ow if ow else 1
            off = l - ol
            if rw >= 0.8:
                set_abs(sh, T, nl + off / ow * Wn, rw * Wn)
            elif abs((l + w / 2) - (ol + ow / 2)) <= 0.1 * ow:
                set_abs(sh, T, nl + (Wn - w) / 2)
            else:
                set_abs(sh, T, nl + off)
            if dep > 0:
                touched[id(sh._element)] = (sh, dep)
    # 구분선: 열 경계에 N-1 개
    seps = list(det['seps'])
    while len(seps) > N - 1:
        s = seps.pop()
        parents.append(s[0]._element.getparent())
        delete_shape(s[0])
    while len(seps) < N - 1 and seps:
        ns, _ = clone(slide, seps[-1][0])
        seps.append((ns,) + tuple(seps[-1][1:]))
    for i, s in enumerate(seps):
        set_abs(s[0], s[4], L + (i + 1) * Wn)
        parents.append(s[0]._element.getparent())
    # 묶음 안 요소를 옮겼으면 묶음 틀 다시 맞춤(깊은 것부터)
    groups = {}
    for p in parents + [sh._element.getparent() for sh, dep in touched.values()]:
        while p is not None and local(p) == 'grpSp':
            groups[id(p)] = (p, 0)
            p = p.getparent()
    lookup = shape_lookup(slide)
    for gid, (gel, _) in sorted(groups.items(), key=lambda kv: -len(list(kv[1][0].iterancestors()))):
        if id(gel) in lookup:
            g, bb, dep, path, T = lookup[id(gel)]
            fit_group_x(g, T)
    return Wn


def delete_column(slide, SW, SH, column, warns):
    det = detect_columns(slide, SW, SH)
    if not det:
        raise ValueError('열 구조를 못 찾음(같은 가로 범위 요소 묶음이 2열 미만)')
    cols = det['cols']
    if not 1 <= column <= len(cols):
        raise ValueError('열 번호 범위 밖: %d (열 %d개)' % (column, len(cols)))
    if len(cols) <= 1:
        raise ValueError('열이 하나뿐이라 지울 수 없음')
    order, dirty = [], []
    for i, c in enumerate(cols):
        if i == column - 1:
            for u in c['units']:
                dirty.append(u[0]._element.getparent())
                delete_shape(u[0])
            continue
        order.append((c['l'], c['r'] - c['l'], c['units']))
    _drop_empty_groups(slide)
    return _relayout(slide, order, det, warns, dirty), len(order)


def duplicate_column(slide, SW, SH, column, warns, max_cols=5):
    det = detect_columns(slide, SW, SH)
    if not det:
        raise ValueError('열 구조를 못 찾음')
    cols = det['cols']
    if len(cols) + 1 > max_cols:
        raise ValueError('열은 %d개까지만 허용(지금 %d개)' % (max_cols, len(cols)))
    if not 1 <= column <= len(cols):
        raise ValueError('열 번호 범위 밖: %d' % column)
    elmap = {}
    new_units = []
    for sh, bb, dep, path, T in cols[column - 1]['units']:
        ns, m = clone(slide, sh)
        elmap.update(m)
        new_units.append((ns, bb, dep, path, T))
    order = []
    for i, c in enumerate(cols):
        order.append((c['l'], c['r'] - c['l'], c['units']))
        if i == column - 1:
            order.append((c['l'], c['r'] - c['l'], new_units))
    Wn = _relayout(slide, order, det, warns)
    return Wn, len(order), elmap


def _drop_empty_groups(slide):
    changed = True
    while changed:
        changed = False
        for sh, *_ in list(walk(slide.shapes)):
            if is_group(sh) and len(list(sh.shapes)) == 0:
                delete_shape(sh)
                changed = True


# ---------------------------------------------------------------- 간트 격자

def detect_gantt(slide, SW, SH):
    """주차 열(뒤쪽 같은 폭 열 4개 이상)이 있는 표를 간트로 봄. 막대 = 표 안쪽에 놓인 얇고 글 없는 선·도형."""
    for sh, bb, dep, path, T in walk(slide.shapes):
        if dep != 0 or kind(sh) != 'table':
            continue
        tb = sh.table
        ws = [c.width for c in tb.columns]
        n = len(ws)
        if n < 6:
            continue
        j = n - 1
        while j > 0 and abs(ws[j - 1] - ws[-1]) <= 0.03 * ws[-1]:
            j -= 1
        weeks = n - j
        if weeks < 4:
            continue
        x0 = sh.left + sum(ws[:j])
        wk = ws[-1]
        ys = [sh.top]
        for row in tb.rows:
            ys.append(ys[-1] + row.height)
        bars = []
        for s2, bb2, d2, p2, T2 in walk(slide.shapes):
            if s2 is sh or d2 != 0 or is_group(s2):
                continue
            if local(s2._element) not in ('cxnSp', 'sp'):
                continue
            if s2.has_text_frame and s2.text_frame.text.strip():
                continue
            l, t, w, h = bb2
            if h > 0.04 * SH or w < 0.5 * wk:
                continue
            cy = t + h / 2
            if not (ys[0] <= cy <= ys[-1]) or l < x0 - 0.5 * wk or l + w > x0 + weeks * wk + 0.5 * wk:
                continue
            row = next((i for i in range(len(ys) - 1) if ys[i] <= cy < ys[i + 1]), len(ys) - 2)
            bars.append((s2, bb2, int(round((l - x0) / wk)) + 1, int(round((l + w - x0) / wk)), row))
        return {'frame': sh, 'x0': x0, 'wk': wk, 'weeks': weeks, 'bars': bars, 'ys': ys}
    return None


def set_bar(sh, x0, wk, start, end):
    """막대 위치 = 첫 주차 위치 + (주차 - 1) × 주차 폭"""
    bx, ex = int(x0 + (start - 1) * wk), int(x0 + end * wk)
    if local(sh._element) == 'cxnSp':
        sh.begin_x, sh.end_x = bx, ex
    else:
        sh.left, sh.width = bx, ex - bx


# ---------------------------------------------------------------- 표 행 삭제

def delete_table_rows(slide, frame, rows, header_rows=1, gantt=None):
    """병합 단위로만 지움. 남은 행 높이 = (머리 행 뺀 원래 합계) ÷ 남은 행 수.
    gantt 가 주어지면 지운 행의 막대도 지우고, 남은 막대는 새 행 위치로 옮김."""
    tbl = frame.table._tbl
    trs = tbl.tr_lst
    n = len(trs)
    rows = sorted(set(rows))
    if any(r < 0 or r >= n for r in rows):
        raise ValueError('행 번호 범위 밖(표는 %d행, 0부터)' % n)
    if any(r < header_rows for r in rows):
        raise ValueError('머리 행(%d행)은 못 지움' % header_rows)
    dele = set(rows)
    for r in range(n):
        for c, tc in enumerate(trs[r].tc_lst):
            span = int(tc.get('rowSpan', '1'))
            if span > 1 and r in dele and not all((r + i) in dele for i in range(span)):
                raise ValueError('행 %d 는 세로 병합 셀의 일부라 못 지움(병합 범위 전체를 함께 지워야 함)' % r)
            if tc.get('vMerge') == '1' and r in dele:
                o = r - 1
                while o >= 0 and trs[o].tc_lst[c].get('vMerge') == '1':
                    o -= 1
                span_o = int(trs[o].tc_lst[c].get('rowSpan', '1'))
                if not all((o + i) in dele for i in range(span_o)):
                    raise ValueError('행 %d 는 세로 병합 범위 일부라 못 지움' % r)
    old_h = [int(t.get('h')) for t in trs]
    old_ys = [frame.top]
    for h in old_h:
        old_ys.append(old_ys[-1] + h)
    for r in rows:
        tbl.remove(trs[r])
    trs = tbl.tr_lst
    head = sum(old_h[:header_rows])
    body_total = sum(old_h) - head
    each = body_total // (len(trs) - header_rows)
    for t in trs[header_rows:]:
        t.set('h', str(each))
    frame.height = head + each * (len(trs) - header_rows)
    removed_bars = 0
    if gantt:
        keep_rows = [r for r in range(n) if r not in dele]
        new_row = {r: i for i, r in enumerate(keep_rows)}
        new_ys = [frame.top]
        for t in trs:
            new_ys.append(new_ys[-1] + int(t.get('h')))
        for s, bb, st, en, row in gantt['bars']:
            if row in dele:
                delete_shape(s)
                removed_bars += 1
                continue
            l, t, w, h = bb
            cy = t + h / 2
            frac = (cy - old_ys[row]) / max(1, old_ys[row + 1] - old_ys[row])
            ncy = new_ys[new_row[row]] + frac * (new_ys[new_row[row] + 1] - new_ys[new_row[row]])
            if local(s._element) == 'cxnSp':
                dy = int(ncy - cy)
                s.begin_y, s.end_y = s.begin_y + dy, s.end_y + dy
            else:
                s.top = int(ncy - h / 2)
    return len(trs), each, removed_bars
