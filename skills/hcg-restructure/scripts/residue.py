"""이전 고객명 잔존 검사: 글 · 표 · 묶음/도형 이름 · 그림 설명 · 발표자 메모 · 마스터/레이아웃 · 문서 속성 · 차트 안 글.

사용:  python residue.py <pptx> --names 이전고객명1 이전고객명2 ... [--names-file client_names.txt] [--slide N] [--json out.json]
  --names-file : 한 줄에 하나씩 적은 고객명 목록(예: hcg-check/references/client_names.txt). --names 와 함께 쓸 수 있음.
  찾으면 종료 코드 1, 없으면 0.
"""
import argparse
import json
import re
import sys

from lxml import etree
from pptx import Presentation

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'


def _texts(root):
    """요소 안의 (종류, 글) 목록. a:p 는 run 을 이어 붙여 한 문단으로(이름이 run 에 쪼개져 있어도 찾음)."""
    for p in root.iter('{%s}p' % A):
        t = ''.join(x.text or '' for x in p.iter('{%s}t' % A))
        if t.strip():
            yield '글', t
    for el in root.iter():
        if isinstance(el.tag, str) and el.tag.endswith('}cNvPr'):
            for at, lab in (('name', '도형·묶음 이름'), ('descr', '그림 설명'), ('title', '도형 제목')):
                if el.get(at):
                    yield lab, el.get(at)
    for v in root.iter('{%s}v' % C):
        if v.text:
            yield '차트 안 글', v.text


def _hit(text, names):
    low = text.lower()
    return [n for n in names if n.lower() in low]


def scan(prs, names, only_slide=None):
    found = []

    def add(where, kind, text, hits):
        for h in hits:
            found.append({'위치': where, '종류': kind, '이름': h, '글': text[:80]})

    for i, s in enumerate(prs.slides, 1):
        if only_slide and i != only_slide:
            continue
        for kind, t in _texts(s._element):
            add('장표 %d' % i, kind, t, _hit(t, names))
        # 차트 부품
        for rel in s.part.rels.values():
            if 'chart' in rel.reltype and not rel.is_external:
                root = etree.fromstring(rel.target_part.blob)
                for kind, t in _texts(root):
                    add('장표 %d' % i, kind, t, _hit(t, names))
        if s.has_notes_slide:
            for kind, t in _texts(s.notes_slide._element):
                add('장표 %d' % i, '발표자 메모', t, _hit(t, names))
        if s.slide_layout is not None:
            for kind, t in _texts(s.slide_layout._element):
                add('장표 %d 레이아웃(%s)' % (i, s.slide_layout.name), '레이아웃 ' + kind, t, _hit(t, names))
    if not only_slide:
        for m in prs.slide_masters:
            for kind, t in _texts(m._element):
                add('슬라이드 마스터', kind, t, _hit(t, names))
        cp = prs.core_properties
        for at, lab in (('title', '제목'), ('subject', '주제'), ('author', '작성자'), ('last_modified_by', '마지막 수정자'),
                        ('keywords', '키워드'), ('comments', '설명'), ('category', '범주')):
            v = getattr(cp, at, None)
            if v:
                add('문서 속성', lab, v, _hit(v, names))
        # 회사명(app.xml)
        for rel in prs.part.package.iter_parts():
            if str(rel.partname).endswith('docProps/app.xml'):
                root = etree.fromstring(rel.blob)
                for el in root.iter():
                    if isinstance(el.tag, str) and el.tag.endswith('}Company') and el.text:
                        add('문서 속성', '회사', el.text, _hit(el.text, names))
    return found


def load_names(args_names, names_file):
    names = list(args_names or [])
    if names_file:
        for ln in open(names_file, encoding='utf-8'):
            ln = ln.strip()
            if ln and not ln.startswith('#'):
                names.append(ln)
    seen, out = set(), []
    for n in names:
        if n and n.lower() not in seen:
            seen.add(n.lower())
            out.append(n)
    return out


def main():
    ap = argparse.ArgumentParser(description='이전 고객명 잔존 검사')
    ap.add_argument('pptx')
    ap.add_argument('--names', nargs='*', default=[], help='찾을 이전 고객명')
    ap.add_argument('--names-file', help='고객명 목록 파일(한 줄에 하나)')
    ap.add_argument('--slide', type=int, help='이 장표만 검사(1부터)')
    ap.add_argument('--json', help='결과를 JSON 으로 저장')
    a = ap.parse_args()
    names = load_names(a.names, a.names_file)
    if not names:
        sys.exit('찾을 이름이 없음: --names 또는 --names-file 을 주세요')
    prs = Presentation(a.pptx)
    found = scan(prs, names, a.slide)
    if a.json:
        json.dump(found, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    if not found:
        print('잔존 없음 (찾은 이름 %d개 기준)' % len(names))
        return 0
    print('잔존 %d곳' % len(found))
    for f in found:
        print('  [%s] %s · 「%s」 · %s' % (f['위치'], f['종류'], f['이름'], f['글']))
    return 1


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.exit(main())
