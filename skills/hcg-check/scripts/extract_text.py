"""모든 장의 글 뽑기: 제목 · 핵심 문장 · 칸별 글 · 표 · 그림 설명 · 발표자 메모 · 레이아웃 이름 · 문서 속성.

사용:  python extract_text.py <pptx> --out text.md      (사람이 읽는 목록)
       python extract_text.py <pptx> --out text.json    (check.py 와 Claude 가 쓰는 구조)
check.py 가 이 파일의 extract() 를 가져다 씀.
"""
import argparse
import json
import re
import sys

from lxml import etree
from pptx import Presentation

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
HEAD_ZONE = 0.28   # 장표 위쪽 28% 안의 글을 머리(제목·핵심 문장) 후보로 봄


def _local(el):
    return etree.QName(el).localname


def _is_group(sh):
    return sh._element.tag == '{%s}grpSp' % NS_P


def _walk(shapes, T=(1.0, 0.0, 1.0, 0.0), depth=0):
    for sh in shapes:
        sx, ox, sy, oy = T
        try:
            l, t = ox + sh.left * sx, oy + sh.top * sy
            w, h = sh.width * sx, sh.height * sy
        except Exception:
            l = t = w = h = 0.0
        yield sh, (l, t, w, h), depth
        if _is_group(sh):
            x = sh._element.grpSpPr.xfrm
            if x is None:
                continue
            ce, co = x.chExt, x.chOff
            nsx = (w / ce.cx) if ce.cx else 1.0
            nsy = (h / ce.cy) if ce.cy else 1.0
            yield from _walk(sh.shapes, (nsx, l - co.x * nsx, nsy, t - co.y * nsy), depth + 1)


def _font_pt(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            if r.font.size:
                return round(r.font.size.pt, 1)
    return None


def _ph_type(sh):
    try:
        if sh.is_placeholder:
            return str(sh.placeholder_format.type).split(' ')[0].split('.')[-1]
    except Exception:
        pass
    return None


def _para_texts(tf):
    out = []
    for p in tf.paragraphs:
        t = ''.join(r.text for r in p.runs).replace('\x0b', '\n')
        if t.strip():
            out.append(t)
    return out


def _slide_info(prs, idx, slide):
    SW, SH = prs.slide_width, prs.slide_height
    info = {'no': idx, 'hidden': slide._element.get('show') == '0',
            'layout': slide.slide_layout.name, 'items': [], 'tables': [], 'pictures': [],
            'counts': {'picture': 0, 'table': 0, 'chart': 0, 'shape': 0, 'group': 0, 'object': 0}}
    title = message = None
    for sh, (l, t, w, h), dep in _walk(slide.shapes):
        tag = _local(sh._element)
        top, left = t / SH, l / SW
        if tag == 'grpSp':
            info['counts']['group'] += 1
            continue
        if tag == 'pic':
            info['counts']['picture'] += 1
            try:
                d = sh._element.nvPicPr.cNvPr.get('descr') or ''
            except Exception:
                d = ''
            info['pictures'].append({'name': sh.name, 'descr': d, 'w_pct': round(w / SW, 3), 'h_pct': round(h / SH, 3),
                                     'top_pct': round(top, 3), 'left_pct': round(left, 3)})
            continue
        if tag == 'graphicFrame':
            if getattr(sh, 'has_table', False) and sh.has_table:
                info['counts']['table'] += 1
                rows = []
                for r in sh.table.rows:
                    rows.append([c.text_frame.text.replace('\x0b', '\n').strip() for c in r.cells])
                info['tables'].append({'name': sh.name, 'rows': rows, 'top_pct': round(top, 3)})
                for ri, r in enumerate(rows):
                    for ci, txt in enumerate(r):
                        if txt:
                            info['items'].append({'kind': 'cell', 'name': '%s r%dc%d' % (sh.name, ri, ci), 'text': txt,
                                                  'top_pct': round(top, 3), 'font_pt': None, 'ph': None})
            elif getattr(sh, 'has_chart', False) and sh.has_chart:
                info['counts']['chart'] += 1
            else:
                info['counts']['object'] += 1
            continue
        if tag == 'cxnSp':
            continue
        # sp
        has_txt = sh.has_text_frame and sh.text_frame.text.strip()
        ph = _ph_type(sh)
        is_box = False
        try:
            is_box = sh._element.nvSpPr.cNvSpPr.get('txBox') == '1'
        except Exception:
            pass
        if not (has_txt and is_box) and not ph:
            info['counts']['shape'] += 1
        if has_txt:
            txt = '\n'.join(_para_texts(sh.text_frame))
            if not txt.strip():
                continue
            it = {'kind': 'text', 'name': sh.name, 'text': txt, 'top_pct': round(top, 3),
                  'left_pct': round(left, 3), 'font_pt': _font_pt(sh.text_frame), 'ph': ph, 'shape_id': sh.shape_id}
            info['items'].append(it)
            if ph and 'TITLE' in ph and title is None:
                title = txt
            elif ph and ph.startswith('BODY') and message is None and top < HEAD_ZONE:
                message = txt
    # 제목 자리표시자가 없는 장표: 위쪽 글 중 짧은 것 = 제목, 가장 긴 것 = 핵심 문장
    if title is None or message is None:
        head = sorted([i for i in info['items'] if i['kind'] == 'text' and i['top_pct'] < HEAD_ZONE],
                      key=lambda i: i['top_pct'])
        if head:
            if title is None:
                title = head[0]['text']
            if message is None:
                rest = [i for i in head if i['text'] != title and len(i['text']) >= 12]
                if rest:
                    message = max(rest, key=lambda i: len(i['text']))['text']
    info['title'] = (title or '').replace('\n', ' ').strip()
    info['message'] = (message or '').replace('\n', ' ').strip()
    notes = ''
    if slide.has_notes_slide:
        notes = '\n'.join(_para_texts(slide.notes_slide.notes_text_frame)) if slide.notes_slide.notes_text_frame else ''
    info['notes'] = notes
    info['all_text'] = '\n'.join(i['text'] for i in info['items'])
    return info


def extract(path):
    prs = Presentation(path)
    cp = prs.core_properties
    d = {'file': path, 'slide_count': len(prs.slides), 'slide_w': prs.slide_width, 'slide_h': prs.slide_height,
         'core': {k: (getattr(cp, k) or '') for k in ('title', 'subject', 'author', 'last_modified_by',
                                                      'keywords', 'comments', 'category')},
         'slides': []}
    for i, s in enumerate(prs.slides, 1):
        d['slides'].append(_slide_info(prs, i, s))
    # 마스터·레이아웃에 있는 글(저작권 문구·로고 이름 등)
    mt = []
    for m in prs.slide_masters:
        for sh, bb, dep in _walk(m.shapes):
            if sh.has_text_frame and sh.text_frame.text.strip():
                mt.append({'where': '마스터', 'name': sh.name, 'text': sh.text_frame.text.strip()})
            elif _local(sh._element) == 'pic':
                mt.append({'where': '마스터', 'name': sh.name, 'text': '(그림) ' + (sh._element.nvPicPr.cNvPr.get('descr') or '')})
        for lo in m.slide_layouts:
            for sh, bb, dep in _walk(lo.shapes):
                if sh.has_text_frame and sh.text_frame.text.strip():
                    mt.append({'where': '레이아웃 ' + lo.name, 'name': sh.name, 'text': sh.text_frame.text.strip()})
                elif _local(sh._element) == 'pic':
                    mt.append({'where': '레이아웃 ' + lo.name, 'name': sh.name,
                               'text': '(그림) ' + (sh._element.nvPicPr.cNvPr.get('descr') or '')})
    d['master_texts'] = mt
    company = ''
    for part in prs.part.package.iter_parts():
        if str(part.partname).endswith('docProps/app.xml'):
            m = re.search(r'<Company>(.*?)</Company>', part.blob.decode('utf-8', 'ignore'), re.S)
            company = m.group(1) if m else ''
    d['core']['company'] = company
    return d


def to_markdown(d):
    out = ['# %s' % d['file'], '장표 %d장 · 문서 속성: %s' % (d['slide_count'], json.dumps(d['core'], ensure_ascii=False)), '']
    for s in d['slides']:
        out.append('## 장표 %d · %s%s · 제목: %s' % (s['no'], s['layout'], ' · 숨김' if s['hidden'] else '', s['title']))
        if s['message']:
            out.append('핵심 문장: ' + s['message'])
        c = s['counts']
        out.append('(그림 %d · 표 %d · 차트 %d · 도형 %d)' % (c['picture'], c['table'], c['chart'], c['shape']))
        for it in s['items']:
            out.append('- [%s] %s' % (it['name'], it['text'].replace('\n', ' / ')))
        for p in s['pictures']:
            if p['descr']:
                out.append('- (그림 설명) %s' % p['descr'])
        if s['notes']:
            out.append('- (발표자 메모) %s' % s['notes'].replace('\n', ' / '))
        out.append('')
    return '\n'.join(out)


def main():
    ap = argparse.ArgumentParser(description='모든 장의 글 뽑기')
    ap.add_argument('pptx')
    ap.add_argument('--out', required=True, help='.md 또는 .json')
    a = ap.parse_args()
    d = extract(a.pptx)
    with open(a.out, 'w', encoding='utf-8') as f:
        if a.out.lower().endswith('.json'):
            json.dump(d, f, ensure_ascii=False, indent=1)
        else:
            f.write(to_markdown(d))
    print('장표 %d장 → %s' % (d['slide_count'], a.out))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
