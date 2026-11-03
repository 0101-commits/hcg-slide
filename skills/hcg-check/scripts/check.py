"""HCG 제안서 점검 28항목(자동 16 · 반자동 12). 자동은 스크립트가 판정, 반자동은 후보만 뽑음(문장 판단은 Claude).

사용:
  python check.py <pptx> --out report.md [--client 새고객명] [--old-names client_names.txt]
        [--figures company_figures.json] [--public rules.json] [--terms terms.json]
        [--requests requests.json] [--rfp-keywords "a,b"] [--meeting meeting.json]
        [--pages 40] [--rfp-weeks 8-12] [--json findings.json]
원본은 읽기만 하고 고치지 않음. 입력이 없는 항목은 「점검 현황」에 「입력 없음」으로 표시하고 건너뜀.
"""
import argparse
import datetime
import json
import os
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract_text import extract  # noqa: E402

FIX, CONFIRM, REF = '고칠 것', '확인할 것', '참고'

# 번호: (이름, 자동/반자동, 무리)
ITEMS = {
    1: ('이전 고객명·다른 회사명 잔재', '자동', '글자 규칙'),
    2: ('자리표시자·작업 메모', '자동', '글자 규칙'),
    3: ('회사 소개 수치', '자동', '글자 규칙'),
    4: ('예시·가설 표시', '자동', '글자 규칙'),
    5: ('익명 표기 통일', '자동', '글자 규칙'),
    6: ('출처·기준일', '반자동', '글자 규칙'),
    7: ('제목 번호·접두', '자동', '장표 구조'),
    8: ('목차 간지', '자동', '장표 구조'),
    9: ('레이아웃 이름', '자동', '장표 구조'),
    10: ('일정 숫자 동기화', '자동', '장표 구조'),
    11: ('분량', '반자동', '장표 구조'),
    12: ('도식 많음·없음', '반자동', '장표 구조'),
    13: ('[참고] 분리', '반자동', '장표 구조'),
    14: ('핵심 문장이 결론형인지', '반자동', '문장 판단'),
    15: ('단정·방어 유발 문구', '반자동', '문장 판단'),
    16: ('Why HCG 주장형', '반자동', '문장 판단'),
    17: ('종결 어미', '자동', '문장 판단'),
    18: ('요청 항목 반영', '반자동', '입력 대조'),
    19: ('RFP 핵심어', '반자동', '입력 대조'),
    20: ('고객 표기', '자동', '입력 대조'),
    21: ('투입 인력 주요 고객', '반자동', '입력 대조'),
    22: ('가격 분리', '자동', '입력 대조'),
    23: ('마지막 회의 지시', '반자동', '입력 대조'),
    24: ('무기명 위반', '자동', '공공·입찰형'),
    25: ('수행 실적 기간', '자동', '공공·입찰형'),
    26: ('지정 목차·필수 장표', '자동', '공공·입찰형'),
    27: ('지정 표지·서식', '자동', '공공·입찰형'),
    28: ('쪽수(본문·요약본)', '반자동', '공공·입찰형'),
}

STD_LAYOUTS = ['표지', '목차', '본문', 'EOD', 'End of document', '끝장', '간지']
LAYOUT_FIX = {'분문': '본문', '본문1__HCG template': '본문', '목차_HCG template': '목차'}


class Ctx:
    def __init__(self, d, a):
        self.d, self.a = d, a
        self.slides = d['slides']
        self.f = []
        self.status = {}
        self.client = a.client
        self.public = json.load(open(a.public, encoding='utf-8')) if a.public else None
        self.figures = json.load(open(a.figures, encoding='utf-8')) if a.figures else None
        self.today = datetime.date.today()
        self.kinds = [self.kind(s, i) for i, s in enumerate(self.slides)]

    def add(self, item, level, slide, where, msg, suggest='', **extra):
        self.f.append(dict({'item': item, 'level': level, 'slide': slide, 'where': where, 'msg': msg, 'suggest': suggest}, **extra))

    def kind(self, s, i):
        lay, t = s['layout'], s['title']
        n = len(self.slides)
        if i == 0 or '표지' in lay:
            return 'cover'
        if i == n - 1 and re.search(r'EOD|End of|끝장', lay + t, re.I):
            return 'eod'
        if re.search(r'EOD|End of', lay, re.I):
            return 'eod'
        if '목차' in lay or re.fullmatch(r'\s*(목차|contents|agenda|table of contents)\s*', t, re.I):
            return 'toc'
        return 'body'


def texts_of(s, with_notes=False):
    for it in s['items']:
        yield it['name'], it['text']
    if with_notes and s['notes']:
        yield '발표자 메모', s['notes']


def ctx_snip(text, m, n=18):
    return text[max(0, m.start() - n):m.end() + n].replace('\n', ' ')


# --------------------------------------------------------------- 자동 16

FOLD_NAMES = 5   # 한 장표에 사전 고객명이 이 수 이상이면 장표당 한 건(확인할 것)으로 묶음


def c1_old_names(c):
    if not c.a.old_names:
        c.status[1] = '입력 없음(--old-names)'
        return
    names = []
    for ln in open(c.a.old_names, encoding='utf-8'):
        ln = ln.strip()
        if ln and not ln.startswith('#'):
            names.append(ln)
    keep = c.client or ''
    names = [n for n in names if not (keep and (n in keep or keep in n))]
    for s in c.slides:
        cases = re.search(r'\[참고\]|Back-?up|사례|주요\s*고객|Clients?|고객사|레퍼런스|실적', s['title'] + s['message'], re.I)
        hits = []   # (칸, 사전 이름, 앞뒤 글)
        for nm, tx in list(texts_of(s, True)) + [('그림 설명', p['descr']) for p in s['pictures'] if p['descr']]:
            low = tx.lower()
            for n in names:
                nl = n.lower()
                if re.fullmatch(r'[a-z0-9 .&\-]+', nl):
                    hit = re.search(r'(?<![a-z0-9])' + re.escape(nl) + r'(?![a-z0-9])', low)
                else:
                    hit = nl in low
                if hit:
                    m = re.search(re.escape(nl), low)
                    snip = tx[max(0, m.start() - 12):m.end() + 12].replace('\n', ' ') if m else ''
                    hits.append((nm, n, snip))
        distinct = []
        for nm, n, snip in hits:
            if n not in distinct:
                distinct.append(n)
        if len(distinct) >= FOLD_NAMES:
            # 사전 이름이 많은 장표는 주요 고객사·실적·인력 목록일 가능성이 커서 장표당 한 건으로 묶음
            c.add(1, CONFIRM, s['no'], '', '고객명 %d개(실적 목록일 수 있음): %s%s' % (len(distinct), ', '.join(distinct[:6]), '…' if len(distinct) > 6 else ''),
                  '실적 목록이면 그대로, 이전 제안의 잔재면 바꿈', names=distinct)
            continue
        for nm, n, snip in hits:
            if cases:
                c.add(1, CONFIRM, s['no'], nm, '이전 고객명 「%s」 — 사례·고객 목록 장표라 의도된 것일 수 있음: …%s…' % (n, snip),
                      '사례면 A사처럼 익명 표기, 새 고객 장표라면 현 고객명으로')
            else:
                c.add(1, FIX, s['no'], nm, '이전 고객명 「%s」 잔존: …%s…' % (n, snip), '현 고객명으로 바꿈')
    cp = c.d['core']
    for k, v in list(cp.items()):
        for n in names:
            if v and n.lower() in v.lower():
                c.add(1, FIX, None, '문서 속성 ' + k, '문서 속성에 이전 고객명 「%s」: %s' % (n, v), '속성 값 정리')
    c.status[1] = '실행'


PLACEHOLDERS = [
    (r'(?<![\d,.])0{3}(?![\d,])', '「000」'), (r'O{3,}|○{3}', '「OOO」'),
    (r'M{2}\.?\s*D{2}', '「MM. DD」'), (r'Y{4}\.?\s*M{2}', '「YYYY. MM」'), (r'(?<![A-Za-z])D{2}(?![A-Za-z])', '「DD」'),
    (r'(?<![\d.])0\s*주(?!\d)', '「0주」'), (r'총\s*0\s*개월', '「총 0개월」'),
    (r'수정\s*중|수정\s*필요|작성\s*중|검토\s*중|TBD|홍길동', '작업 메모'),
    (r'\[[^\]]*Message[^\]]*\]', '「[Message 필요]」'),
    (r'소목차[^.\n]{0,20}(없는|필요|기입)|기입해\s*주세요|내용을\s*입력|클릭하여|Template Guide', '틀 안내문'),
    (r'(?:^|\s)(?:MH|HE|CK|YH|ds)\s*(?:수정|확인|작성|검토|추가|중|님)', '작성자 이니셜 메모'),
]


def c2_placeholders(c):
    anon = bool(c.public and c.public.get('무기명'))
    for s in c.slides:
        for nm, tx in texts_of(s, True):
            for pat, lab in PLACEHOLDERS:
                for m in re.finditer(pat, tx):
                    if '인력' in s['title'] + nm and re.fullmatch(r'O{3,}|○{3}', m.group()) and anon:
                        continue
                    lvl = FIX
                    sug = '삭제하거나 확정값 입력'
                    if lab in ('「OOO」', '「000」') and re.search(r'인력|투입|견적|가격|비용|만원', s['title'] + tx):
                        lvl, sug = CONFIRM, '인력 가림·가격 미정 표시일 수 있음 — 의도한 것인지 확인'
                    c.add(2, lvl, s['no'], nm, '%s 잔존: …%s…' % (lab, ctx_snip(tx, m)), sug)
    cover = c.slides[0] if c.slides else None
    if cover:
        for nm, tx in texts_of(cover):
            m = re.search(r'(20\d\d)\s*[.\-년]\s*(\d{1,2})', tx)
            if m:
                y, mo = int(m.group(1)), int(m.group(2))
                if 1 <= mo <= 12 and (c.today.year * 12 + c.today.month) - (y * 12 + mo) > 12:
                    c.add(2, CONFIRM, 1, nm, '표지 날짜 %d.%02d 가 1년 넘게 지남(예전 날짜 잔존?)' % (y, mo), '제출일로 갱신')
    c.status[2] = '실행'


MONTH_TOK = re.compile(r'(20\d\d)\s*(?:[.\-]\s*(\d{1,2})|년\s*(\d{1,2})\s*월?)')


def _month_near(t, m):
    """수치 문구가 든 줄에 붙은 기준월(예: 2026.01 · 2025년 1월 · (2025.01 기준)) → 'YYYY.MM' 또는 None"""
    a = t.rfind('\n', 0, m.start()) + 1
    b = t.find('\n', m.end())
    for mm in MONTH_TOK.finditer(t[a:(len(t) if b < 0 else b)]):
        mo = int(mm.group(2) or mm.group(3))
        if 1 <= mo <= 12:
            return '%s.%02d' % (mm.group(1), mo)
    return None


def _figs(c):
    """장표별 (고객 수, HR 인력, 서비스 연수, 기준월) 값 찾기"""
    return [x[:4] for x in _figs_ex(c)]


def _figs_ex(c):
    """_figs + (읽은 그대로의 글자, 붙은 기준월)"""
    out = []
    pc = [r'고객(?:사|기업)?(?:\s*수)?[^\d\n]{0,8}?(\d[\d,]{2,5})\s*(?:여\s*)?(?:개\s*사|개|사|곳|社)',
          r'(\d[\d,]{2,5})\s*(?:여\s*)?(?:개\s*사|개|사|곳|社)\s*(?:의\s*)?(?:고객|이상의 고객|고객사)']
    pp = [r'(?:HR\s*)?(?:전문|컨설팅)?\s*인력[^\d\n]{0,8}?(\d{2,3})\s*명', r'(\d{2,3})\s*(?:여\s*)?명(?:의)?\s*(?:HR\s*)?(?:전문\s*)?인력',
          r'(?:HR\s*)?전문가[^\d\n]{0,6}(\d{2,3})\s*명']
    py = [r'(?:서비스|업력|경험)[^\d\n]{0,6}(\d{2})\s*년', r'(\d{2})\s*년(?:의|간의)?\s*(?:서비스|업력|경험)']
    pm = r'(20\d\d)\s*[.\-년]\s*(\d{1,2})\s*(?:월)?\s*(?:기준|현재)'
    for s in c.slides:
        t = s['all_text']
        for kind, pats in (('고객수', pc), ('HR인력', pp), ('서비스연수', py)):
            for p in pats:
                for m in re.finditer(p, t):
                    out.append((s['no'], kind, int(m.group(1).replace(',', '')), ctx_snip(t, m), m.group(1), _month_near(t, m)))
        for m in re.finditer(pm, t):
            out.append((s['no'], '기준월', '%s.%02d' % (m.group(1), int(m.group(2))), ctx_snip(t, m), m.group(0), None))
    return out


def c3_company_figures(c):
    found = _figs_ex(c)
    byk = defaultdict(set)
    for no, k, v, snip, raw, mon in found:
        byk[k].add(v)
    for k in ('고객수', 'HR인력', '서비스연수'):
        if len(byk[k]) > 1:
            where = '; '.join('장표 %d: %s' % (no, v) for no, kk, v, sn, raw, mon in found if kk == k)
            c.add(3, FIX, None, k, '%s 값이 장표마다 다름 (%s)' % (k, where), '한 값으로 통일(수치 표 기준)')
    if c.figures:
        ok = {'고객수': set(), 'HR인력': set()}
        months = set()
        for r in c.figures.get('기준', []):
            if r.get('기준월') in c.figures.get('현재_허용', []):
                ok['고객수'].add(r['고객수'])
                ok['HR인력'].add(r['HR인력'])
                months.add(r['기준월'])
        for no, k, v, snip, raw, mon in found:
            if k in ok and ok[k] and v not in ok[k]:
                old = [r['기준월'] for r in c.figures['기준'] if r.get({'고객수': '고객수', 'HR인력': 'HR인력'}[k]) == v]
                oldtxt = (' · 옛 기준 ' + ','.join(old)) if old else ''
                sug = '수치 표 값(%s)으로 갱신, 기준월 함께 표기 — ※표는 관리자 확인 필요' % '/'.join(str(x) for x in sorted(ok[k]))
                if mon:
                    # 기준월이 붙은 수치: 공식 기준월과 다르면 옛 기준 값이라 고칠 것
                    c.add(3, FIX, no, k, '%s %s 가 기준월 %s 로 표기돼 공식 기준월(%s)과 다름(…%s…)%s' % (k, v, mon, ', '.join(sorted(months)), snip, oldtxt), sug)
                else:
                    c.add(3, CONFIRM, no, k, '%s %s 가 수치 표와 다름 — 기준월 표기가 없어 옛 제안의 당시 값일 수 있음(…%s…)%s' % (k, v, snip, oldtxt), sug)
            if k == '기준월' and months and v not in months:
                c.add(3, CONFIRM, no, '기준월', '기준월 %s 가 수치 표(%s)와 다름' % (v, ', '.join(sorted(months))), '기준월 갱신')
        c.status[3] = '실행(수치 표 대조)' + (' · 표는 관리자 확인 필요' if c.figures.get('관리자_확인_필요') else '')
    else:
        c.status[3] = '일부 실행(장표끼리 비교만, --figures 없음)'


def c4_example_labels(c):
    for s in c.slides:
        for nm, tx in texts_of(s):
            for m in re.finditer(r'Ex\)|E\.g\.?|Illustrative(?: use only)?|예\)\s', tx):
                c.add(4, FIX, s['no'], nm, '영문·옛 표기 「%s」: …%s…' % (m.group().strip(), ctx_snip(tx, m)), '「예시적」·「가설적」으로 통일')
        head = s['title'] + ' ' + s['message']
        if re.search(r'시뮬레이션|Simulation|Output Image|산출물 이미지|Payout|Pay Band', head, re.I) and \
                not re.search(r'예시|가설|Illustrative', s['all_text'] + head) and (s['counts']['table'] or s['counts']['shape'] > 5):
            c.add(4, CONFIRM, s['no'], s['title'][:20], '가정 수치·산출물 예시로 보이는데 「예시적」·「가설적」 표시가 없음', '제목이나 도식에 「예시적」 표시')
    c.status[4] = '실행'


def c5_anonymous(c):
    forms = Counter()
    where = defaultdict(list)
    for s in c.slides:
        for nm, tx in texts_of(s):
            for m in re.finditer(r'(?<![A-Za-z가-힣])([A-Z])\s?(사|社)(?![가-힣]{2})', tx):
                forms[m.group(2)] += 1
                where[m.group(2)].append(s['no'])
    if len(forms) > 1:
        c.add(5, FIX, None, '익명 표기', '「A사」와 「A社」가 섞여 있음 (사 %d곳: 장표 %s / 社 %d곳: 장표 %s)' % (
            forms['사'], sorted(set(where['사']))[:8], forms['社'], sorted(set(where['社']))[:8]), '한 형식(A사)으로')
    c.status[5] = '실행(고객 실명 사례는 1번 항목이 함께 잡음)'


NUM_PREFIX = re.compile(r'^\s*((?:[IVX]+|[Ⅰ-Ⅹ])\.|\d+(?:\.\d+)*\.?|\d+\))\s*')


def c7_title_prefix(c):
    seq = []
    for s, k in zip(c.slides, c.kinds):
        if k != 'body':
            continue
        m = NUM_PREFIX.match(s['title'])
        seq.append((s['no'], m.group(1) if m else None))
    if not seq:
        c.status[7] = '실행(본문 장표 없음)'
        return
    styles = Counter()
    for no, p in seq:
        if p is None:
            styles['없음'] += 1
        elif re.match(r'[IVXⅠ-Ⅹ]', p):
            styles['로마자'] += 1
        else:
            styles['숫자'] += 1
    with_p = styles['로마자'] + styles['숫자']
    if with_p and styles['없음'] >= 0.1 * len(seq):
        miss = [no for no, p in seq if p is None]
        c.add(7, FIX, None, '제목 접두', '제목 번호·접두가 있는 장표 %d장, 없는 장표 %d장 (없는 장표: %s)' % (
            with_p, styles['없음'], miss[:12]), '목차 기준으로 접두 일괄 부여')
    if styles['로마자'] and styles['숫자']:
        c.add(7, CONFIRM, None, '제목 접두', '로마자·숫자 접두가 섞임 (로마자 %d · 숫자 %d)' % (styles['로마자'], styles['숫자']), '한 형식으로(대분류 로마자·소분류 숫자면 정상)')
    last = None
    for no, p in seq:
        if p and re.fullmatch(r'\d+\.', p):
            n = int(p[:-1])
            if last is not None and n > last + 1:
                c.add(7, FIX, no, '제목 번호', '번호가 %d → %d 로 건너뜀' % (last, n), '번호 다시 매김')
            last = n
    c.status[7] = '실행'


def c8_section_dividers(c):
    if len(c.slides) < 8:
        c.status[8] = '건너뜀(장표 8장 미만)'
        return
    toc_idx = [i for i, k in enumerate(c.kinds) if k == 'toc']
    if not toc_idx:
        c.add(8, CONFIRM, None, '목차', '목차 장표를 못 찾음(레이아웃 이름 「목차」 · 제목 「목차」 기준)', '목차 장표 확인')
        c.status[8] = '실행(목차 없음)'
        return
    first = c.slides[toc_idx[0]]
    entries = [ln.strip() for ln in first['all_text'].split('\n') if len(ln.strip()) >= 2 and ln.strip().lower() not in ('목차', 'contents')]
    n_div = len(toc_idx) - 1
    n_sec = len([e for e in entries if NUM_PREFIX.match(e) or len(e) <= 30])
    if n_sec >= 3 and n_div < min(n_sec, 3) - 1 and len(c.slides) >= 20:
        c.add(8, FIX, None, '목차 간지', '섹션 후보 %d개인데 목차 간지는 %d장뿐 (본문 %d장)' % (n_sec, len(toc_idx), len(c.slides)),
              '섹션이 바뀔 때마다 목차 간지를 넣고 현재 위치 강조')
    ent_norm = [re.sub(r'\W', '', NUM_PREFIX.sub('', e)) for e in entries]
    for i in toc_idx[1:]:
        s = c.slides[i]
        lines = [re.sub(r'\W', '', NUM_PREFIX.sub('', ln.strip())) for ln in s['all_text'].split('\n') if ln.strip()]
        bad = [ln for ln in lines if ln and ent_norm and not any(ln in e or e in ln for e in ent_norm)]
        if bad:
            c.add(8, CONFIRM, s['no'], '간지 항목', '간지 항목명 %s 이(가) 첫 목차와 다름' % bad[:3], '첫 목차 항목명과 맞춤')
    c.status[8] = '실행'


def c9_layout_names(c):
    allowed = [x.lower() for x in STD_LAYOUTS]
    seen = defaultdict(list)
    for s in c.slides:
        base = re.sub(r'^\d+_', '', s['layout']).strip().lower()
        if base not in allowed:
            seen[s['layout']].append(s['no'])
    for lay, nos in seen.items():
        c.add(9, FIX, nos[0], '레이아웃 「%s」' % lay, '표준 밖 레이아웃 이름 「%s」 (장표 %s)' % (lay, nos[:10]),
              '표준 레이아웃 「%s」로 바꿈' % LAYOUT_FIX.get(lay, '표지·목차·본문·끝장 중 하나'))
    c.status[9] = '실행'


def _weeks(c):
    out = []
    for s in c.slides:
        t = s['all_text']
        for p in (r'총\s*(\d{1,2})\s*주', r'(?<![\d~\-W])(\d{1,2})\s*주\s*(?:간|일정|프로젝트|동안|수행)', r'(?<![\d~\-W])(\d{1,2})\s*주\s*[·ㆍ,]\s*보고'):
            for m in re.finditer(p, t):
                out.append((s['no'], int(m.group(1)), ctx_snip(t, m)))
        for tb in s['tables']:
            ws = [int(x) for r in tb['rows'] for x in [re.fullmatch(r'W(\d{1,2})', cc.strip()) and re.fullmatch(r'W(\d{1,2})', cc.strip()).group(1) for cc in r] if x]
            if len(ws) >= 4:
                out.append((s['no'], max(ws), '간트 표 W1~W%d' % max(ws)))
    return out


def c10_weeks(c):
    w = _weeks(c)
    vals = Counter(v for _, v, _ in w)
    if len(vals) > 1:
        c.add(10, FIX, None, '일정 주수', '일정 주수가 장표마다 다름: ' + '; '.join('%d주 ← 장표 %s' % (v, sorted({n for n, vv, _ in w if vv == v})[:8]) for v in vals),
              '기간 값 하나로 목차·전체 흐름·간트·본문 문장을 모두 갱신(일정 옵션 2안이면 의도된 것일 수 있음)')
    if c.a.rfp_weeks and w:
        lo, hi = c.a.rfp_weeks
        for n, v, sn in w:
            if not lo <= v <= hi:
                c.add(10, FIX, n, '일정 주수', '%d주가 RFP 기간 %d~%d주 밖(…%s…)' % (v, lo, hi, sn), 'RFP 기간 안으로')
    c.status[10] = '실행' + ('' if w else '(주수 표기를 못 찾음)')


TAIL_NOUN = re.compile(r'(?:함|임|됨|음|함\)|임\)|있음|없음|것임|예정임|필요함|가능함)[\.\s]*$')
TAIL_DA = re.compile(r'(?:다|한다|이다|된다|있다)\s*\.?$')
TAIL_POLITE = re.compile(r'(?:습니다|입니다|합니다|드립니다|세요|해요|예요|이에요)\s*\.?$')


GREET_TAIL = re.compile(r'(?:감사합니다|감사드립니다|드립니다|안녕하십니까|안녕하세요)\s*[.!]?$')
GREET_SLIDE = re.compile(r'인사말|Greeting|Thank\s*you|감사합니다|감사의\s*말', re.I)


def _in_quote(s):
    """줄 끝이 따옴표 안이면 True(인용문)"""
    inq = False
    for ch in s:
        if ch in '“「『‘':
            inq = True
        elif ch in '”」』’':
            inq = False
        elif ch == '"':
            inq = not inq
    return inq


def c17_endings(c):
    cnt = Counter()
    ex = defaultdict(list)
    for s, k in zip(c.slides, c.kinds):
        if k != 'body' or GREET_SLIDE.search(s['title'] + ' ' + s['message']):
            continue
        cand = [s['message']] if s['message'] else []
        for it in s['items']:
            if it['kind'] == 'text' and it['ph'] is None and len(it['text']) > 25:
                cand += [ln for ln in it['text'].split('\n') if len(ln) > 20]
        for ln in cand:
            ln = ln.strip()
            if TAIL_POLITE.search(ln):
                if GREET_TAIL.search(ln) or _in_quote(ln):
                    continue   # 인사말·인용문은 뺌
                cnt['습니다체'] += 1
                ex['습니다체'].append((s['no'], ln))
            elif TAIL_NOUN.search(ln):
                cnt['~함체'] += 1
                if re.search(r'하겠음$', ln):
                    ex['하겠음'].append((s['no'], ln))
            elif TAIL_DA.search(ln):
                cnt['~다체'] += 1
                ex['~다체'].append((s['no'], ln))
    main = cnt.most_common(1)[0][0] if cnt else None
    total = sum(cnt.values())
    # 목표 체는 「~함」: 습니다체 문장을 확인할 것으로, 문서 전체가 습니다체면 한 건으로 묶음
    if total >= 8 and cnt['습니다체'] >= 0.7 * total:
        c.add(17, CONFIRM, None, '종결 어미', '전체가 습니다체 — ~함 체로 바꿀지 (습니다체 %d문장 / 판정 %d문장)' % (cnt['습니다체'], total),
              '「~함」체로 일괄 변환할지 사람이 정함')
    else:
        for no, ln in ex['습니다체']:
            c.add(17, CONFIRM, no, '종결 어미', '습니다체 문장(목표 「~함」체): …%s' % ln[-24:], '「~함」으로 통일(인사말·인용이면 그대로)')
    if ex['하겠음']:
        c.add(17, CONFIRM, ex['하겠음'][0][0], '종결 어미', '「~하겠음」 %d곳 (장표 %s)' % (len(ex['하겠음']), sorted({n for n, _ in ex['하겠음']})[:8]), '「~함」·「~할 것임」 중 하나로 통일')
    c.status[17] = '실행(주 어미: %s)' % (main or '판정할 문장 없음')


def c20_terms(c):
    if not c.a.terms:
        c.status[20] = '입력 없음(--terms)'
        return
    terms = json.load(open(c.a.terms, encoding='utf-8'))
    for s in c.slides:
        for nm, tx in texts_of(s):
            for bad, good in terms.items():
                for m in re.finditer(re.escape(bad), tx):
                    c.add(20, FIX, s['no'], nm, '「%s」 → 「%s」: …%s…' % (bad, good, ctx_snip(tx, m)), '사전대로 「%s」' % good)
    c.status[20] = '실행'


MONEY = re.compile(r'\d[\d,.]*\s*(?:억|천만|백만|만)\s*원|₩\s*\d|\d[\d,]*\s*원\b|0{3}\s*만\s*원')
PRICE_CTX = re.compile(r'견적|용역비|수행\s*비용|프로젝트\s*비용|부가세|VAT|Fee|M/M|가격|제안\s*금액|산출\s*내역', re.I)


def c22_price(c):
    for s in c.slides:
        t = s['title'] + '\n' + s['all_text']
        ms = list(MONEY.finditer(t))
        if ms and PRICE_CTX.search(t):
            c.add(22, FIX, s['no'], '금액', '본문에 가격·금액 표기: …%s…' % ctx_snip(t, ms[0]), '본문에서 빼고 「본문 밖(가격 별도 파일)」 표시')
        elif ms:
            c.add(22, REF, s['no'], '금액', '금액 표기(고객 재무 수치일 수 있음): …%s…' % ctx_snip(t, ms[0]), '가격이면 본문 밖으로')
    c.status[22] = '실행'


# --------------------------------------------------------------- 공공·입찰형 자동 4 + 무기명

LEAK_WORDS = ['HCG', '휴먼컨설팅그룹', 'Human Consulting', 'humanconsulting', 'talenx', '휴넬', 'hunel', 'JaDE', 'DNapp']
LEAK_RE = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[a-z]{2,}|https?://\S+|www\.\S+|©|저작권|Copyright', re.I)
INFER_RE = re.compile(r'고객\s*(?:사\s*)?\d{3,}\s*(?:개|사|곳)|\d{4}\s*년\s*설립|서비스\s*\d{2}\s*년|전문\s*인력\s*\d{2,3}\s*명|재계약률')


def c24_anonymous(c):
    r = c.public
    if not r or not r.get('무기명'):
        c.status[24] = '꺼짐(발주처 규정 없음)'
        return
    words = LEAK_WORDS + list(r.get('금지어', [])) + list(r.get('금지_인명', []))
    for s in c.slides:
        where = [(nm, tx) for nm, tx in texts_of(s, True)]
        for nm, tx in where:
            for w in words:
                if w.lower() in tx.lower():
                    c.add(24, FIX, s['no'], nm, '무기명 위반 — 금지어 「%s」: …%s…' % (w, tx[max(0, tx.lower().find(w.lower()) - 10):][:40].replace('\n', ' ')), '지우거나 「제안사」·「당사」로(인력은 OOO·등급·인원수)')
            m = LEAK_RE.search(tx)
            if m:
                c.add(24, FIX, s['no'], nm, '무기명 위반 — 주소·메일·저작권 문구: %s' % m.group()[:40], '삭제')
            m = INFER_RE.search(tx)
            if m:
                c.add(24, CONFIRM, s['no'], nm, '회사를 유추할 수 있는 수치·표현 후보: …%s…' % ctx_snip(tx, m), '사람이 정함(지우기·일반화)')
        for p in s['pictures']:
            if re.search(r'logo|로고|워터마크|watermark', p['name'] + ' ' + p['descr'], re.I):
                c.add(24, FIX, s['no'], '그림 ' + p['name'], '로고·워터마크로 보이는 그림 「%s」' % (p['descr'] or p['name']), '삭제(고객사 로고가 아닌지 확인)')
        if s['hidden']:
            c.add(24, CONFIRM, s['no'], '숨긴 장표', '숨긴 장표가 있음 — 제출본에 남아 있으면 안 됨', '삭제')
    for m in c.d['master_texts']:
        for w in words + ['©', '저작권', 'Copyright']:
            if w.lower() in m['text'].lower():
                c.add(24, FIX, None, m['where'], '무기명 위반 — %s의 「%s」 안에 「%s」' % (m['where'], m['name'], w), '마스터·레이아웃에서 지움')
    for k, v in c.d['core'].items():
        for w in words:
            if v and w.lower() in v.lower():
                c.add(24, FIX, None, '문서 속성 ' + k, '문서 속성 %s 에 「%s」: %s' % (k, w, v), '속성 지움')
    c.status[24] = '실행'


def c25_track_record(c):
    r = c.public
    if not r or not r.get('최근_N년'):
        c.status[25] = '꺼짐(발주처 규정 없음)'
        return
    n, base = int(r['최근_N년']), int(r.get('기준연도', c.today.year))
    need = int(r.get('최소_실적건수', 0))
    ok_rows = 0
    for s in c.slides:
        if not re.search(r'실적|수행\s*사례|레퍼런스', s['title'] + s['message']):
            continue
        for tb in s['tables']:
            for ri, row in enumerate(tb['rows']):
                ys = [int(y) for cell in row for y in re.findall(r'(?<!\d)(20\d\d)(?!\d)', cell)]
                if not ys:
                    continue
                y = max(ys)
                if base - y >= n:
                    c.add(25, FIX, s['no'], '실적 표 %d행' % (ri + 1), '%d년 실적은 최근 %d개년(%d~%d) 밖: %s' % (y, n, base - n + 1, base, ' | '.join(row)[:50]),
                          '기간 밖 실적을 빼고 유사 실적으로 채움')
                else:
                    ok_rows += 1
        if not re.search(r'최근\s*%d\s*개?\s*년|최근\s*%d\s*년' % (n, n), s['all_text'] + s['title']):
            c.add(25, CONFIRM, s['no'], '실적 표', '「최근 %d개년 유사 사업」 표기가 없음' % n, '표 제목에 표기')
    if need and ok_rows < need:
        c.add(25, FIX, None, '실적 건수', '기간 안 실적 %d건 < 기준 %d건(표에서 연도가 읽힌 행 기준)' % (ok_rows, need), '유사 실적 추가')
    c.status[25] = '실행'


def c26_required(c):
    r = c.public
    if not r or not (r.get('지정_목차') or r.get('필수_장표')):
        c.status[26] = '꺼짐(발주처 규정 없음)'
        return
    titles = [(s['no'], re.sub(r'\W', '', s['title'] + s['message'])) for s in c.slides]
    last = -1
    for t in r.get('지정_목차', []):
        key = re.sub(r'\W', '', t)
        pos = next((i for i, (no, tt) in enumerate(titles) if key in tt), None)
        if pos is None:
            c.add(26, FIX, None, '지정 목차', '지정 목차 「%s」 에 해당하는 장표 제목이 없음' % t, '문구를 맞추거나 장표 추가')
        elif pos < last:
            c.add(26, FIX, titles[pos][0], '지정 목차', '「%s」 가 지정 순서보다 앞에 있음' % t, '평가 항목 순서대로')
        else:
            last = pos
    for t in r.get('필수_장표', []):
        key = re.sub(r'\W', '', t)
        if not any(key in tt for no, tt in titles):
            c.add(26, FIX, None, '필수 장표', '필수 장표 「%s」 가 없음' % t, '장표 추가')
    c.status[26] = '실행'


def c27_format(c):
    r = c.public
    if not r or not (r.get('지정_레이아웃') or r.get('인력표_열')):
        c.status[27] = '꺼짐(발주처 규정 없음)'
        return
    if r.get('지정_레이아웃'):
        allowed = [x.lower() for x in r['지정_레이아웃']]
        for s in c.slides:
            if s['layout'].lower() not in allowed:
                c.add(27, FIX, s['no'], '레이아웃', '지정 서식 밖 레이아웃 「%s」' % s['layout'], '발주처 지정 양식으로')
    if r.get('인력표_열'):
        need = [re.sub(r'\W', '', x) for x in r['인력표_열']]
        for s in c.slides:
            if re.search(r'투입\s*인력|수행\s*조직|인력\s*(?:구성|현황)', s['title']):
                for tb in s['tables']:
                    hd = [re.sub(r'\W', '', x) for x in (tb['rows'][0] if tb['rows'] else [])]
                    miss = [x for x in need if not any(x in h for h in hd)]
                    if miss:
                        c.add(27, FIX, s['no'], '인력 표 열', '발주처 서식 열 %s 이(가) 없음' % miss, '지정 서식 열로')
    c.status[27] = '실행'


# --------------------------------------------------------------- 반자동 12 (후보만)

def c6_sources(c):
    for s, k in zip(c.slides, c.kinds):
        if k != 'body':
            continue
        t = s['all_text']
        nums = len(re.findall(r'\d[\d,.]*\s*(?:%|억|조|만\s*명|배|CAGR)', t))
        if (s['counts']['chart'] or nums >= 4) and not re.search(r'출처|Source|자료\s*[:：]|기준\)|\(E\)|※', t, re.I):
            c.add(6, CONFIRM, s['no'], s['title'][:20], '수치·차트가 있는데 출처·기준일 각주가 안 보임(수치 %d개, 차트 %d)' % (nums, s['counts']['chart']),
                  '출처·기준일 각주 추가, 추정치는 (E) 표기')
    c.status[6] = '후보(Claude 가 수치 근거 판단)'


def _is_ref(s):
    return bool(re.search(r'\[참고\]|\[Back-?up\]|Back-?up|Appendix|부록|별첨', s['title'] + ' ' + s['message'], re.I))


def c11_volume(c):
    body = [s for s, k in zip(c.slides, c.kinds) if k == 'body']
    main = [s for s in body if not _is_ref(s)]
    refs = [s for s in body if _is_ref(s)]
    flow = [s for s in main if re.search(r'전체\s*흐름|Overall|Approach|프로세스 개요|Roadmap', s['title'] + s['message'], re.I)]
    c.add(11, REF, None, '장수', '전체 %d장 · 본문 %d장(표지·목차·간지·끝장·[참고] 뺌) · [참고]·Back-up %d장' % (len(c.slides), len(main), len(refs)))
    if c.a.pages and len(main) > c.a.pages:
        c.add(11, CONFIRM, None, '분량', '본문 %d장 vs 지시 %d장 (%d장 초과)' % (len(main), c.a.pages, len(main) - c.a.pages),
              '줄일 후보: [참고]로 내릴 장표, 전체 흐름 장표 %d장(%s), 방법 세부·사례' % (len(flow), [s['no'] for s in flow][:6]))
    c.status[11] = '실행' + ('' if c.a.pages else '(분량 지시 없음 → 장수만 표시)')


def c12_diagrams(c):
    textonly, heavy = [], []
    for s, k in zip(c.slides, c.kinds):
        if k != 'body':
            continue
        cn = s['counts']
        vis = cn['picture'] + cn['table'] + cn['chart'] + cn['shape'] + cn['object']
        if vis <= 2 and len(s['all_text']) > 150:
            textonly.append(s['no'])
        if cn['shape'] + cn['picture'] >= 40:
            heavy.append((s['no'], cn['shape'] + cn['picture']))
    if textonly:
        c.add(12, CONFIRM, textonly[0], '글만 있는 장표', '글만 있고 도식·표가 거의 없는 장표 %d장: %s' % (len(textonly), textonly[:15]), '도식 추가 후보')
    for no, n in heavy:
        c.add(12, CONFIRM, no, '도식 과다', '그림·도형 %d개로 많음' % n, '한 장 한 메시지로 단순화 후보')
    c.status[12] = '후보'


def c13_ref_split(c):
    unmarked, marked_idx, body_idx = [], [], []
    for i, (s, k) in enumerate(zip(c.slides, c.kinds)):
        if k != 'body':
            continue
        if _is_ref(s):
            marked_idx.append(i)
        else:
            body_idx.append(i)
            if re.search(r'사례|벤치마킹|Benchmarking|Case\b|Best Practice|레퍼런스|Reference', s['title'] + s['message'], re.I):
                unmarked.append(s['no'])
    if unmarked:
        c.add(13, CONFIRM, unmarked[0], '[참고] 표시', '사례·벤치마킹으로 보이는데 [참고] 표시가 없는 장표 %d장: %s' % (len(unmarked), unmarked[:15]),
              '[참고] 표시를 붙이고 본문 뒤로(핵심 요청 내용이면 본문 유지)')
    if marked_idx and body_idx and min(marked_idx) < max(body_idx):
        early = [c.slides[i]['no'] for i in marked_idx if i < max(body_idx)]
        c.add(13, REF, early[0], '[참고] 위치', '[참고] 장표 %d장이 본문 장표 사이에 있음: %s' % (len(early), early[:12]), '본문 뒤로 옮길지 판단')
    c.status[13] = '후보'


NOUN_END = re.compile(r'(현황|이해|개요|방향|방안|체계|구성|내용|소개|종류|사례|절차|프로세스|일정|목적|범위|필요|결과|구조|모델|제도|항목|기준)\s*$')
VERB_END = re.compile(r'(함|임|음|됨|있음|없음|필요함|시점|것임|예정|가능|요구됨|요청됨)\s*\.?$|다\s*\.?$')


def c14_message(c):
    cand = 0
    for s, k in zip(c.slides, c.kinds):
        if k != 'body':
            continue
        m = s['message']
        if not m:
            if len(s['all_text']) > 80:
                c.add(14, CONFIRM, s['no'], s['title'][:20], '핵심 문장(제목 아래 한 줄)이 없음', '결론형 문장 후보 2개')
                cand += 1
            continue
        if len(m) < 18 or (NOUN_END.search(m) and not VERB_END.search(m)):
            c.add(14, CONFIRM, s['no'], '핵심 문장', '명사 나열·짧은 문장으로 보임: 「%s」' % m[:50], '결론형 한 문장 후보 2개(Claude 가 작성)')
            cand += 1
    c.status[14] = '후보 %d건(Claude 가 문장 판단)' % cand


ASSERT = ['우려', '불안정', '구식', '리스크 높음', '반드시', '최적의', '최적 ', '문제는', '해야 한다', '해야 함', '비효율', '실패', '심각', '승리',
          '적극 활용할 예정', '높은 이해도', '과도한']


def c15_assertive(c):
    for s in c.slides:
        for nm, tx in texts_of(s):
            for w in ASSERT:
                i = tx.find(w)
                if i >= 0:
                    c.add(15, CONFIRM, s['no'], nm, '단정·방어 유발 후보 「%s」: …%s…' % (w.strip(), tx[max(0, i - 15):i + 20].replace('\n', ' ')),
                          '사실 진술형·「~필요한 시점」으로(Claude 가 문맥 판단)')
    c.status[15] = '후보'


CLAIM = re.compile(r'성공적인[^.\n]{0,20}(?:실시|수행|구축)|국내\s*최고|최고|최대|유일|완벽|보장|확실')


def c16_why(c):
    for s in c.slides:
        why = re.search(r'Why\s*(?:HCG|Us)|왜\s*HCG|차별', s['title'] + s['message'], re.I)
        for nm, tx in texts_of(s):
            for m in CLAIM.finditer(tx):
                if why or re.search(r'HCG|당사', tx):
                    c.add(16, CONFIRM, s['no'], nm, '결과 단정·최상급 후보: …%s…' % ctx_snip(tx, m), '「~경험을 보유함」·사례형으로(국내 최고·최대는 근거 수치와 함께)')
    c.status[16] = '후보'


def c18_requests(c):
    if not c.a.requests:
        c.status[18] = '입력 없음(--requests)'
        return
    reqs = json.load(open(c.a.requests, encoding='utf-8'))
    rows = []
    for r in reqs:
        if r.get('본문밖'):
            rows.append((r, [], '본문 밖'))
            continue
        kws = r.get('keywords') or re.findall(r'[가-힣A-Za-z]{2,}', r['항목'])
        hit = []
        for s, k in zip(c.slides, c.kinds):
            if k != 'body':
                continue
            hay = s['title'] + ' ' + s['message']
            if any(kw.lower() in hay.lower() for kw in kws):
                hit.append(s)
            elif any(kw.lower() in s['all_text'].lower() for kw in kws):
                hit.append(dict(s, _weak=True))
        strong = [s for s in hit if not s.get('_weak')]
        main = [s for s in strong if not _is_ref(s)]
        if not hit:
            v = '누락'
        elif len(main) >= 2 or (len(main) == 1 and len(hit) >= 2):
            v = '충분'
        else:
            v = '부분'
        rows.append((r, hit, v))
        if v == '누락':
            c.add(18, CONFIRM, None, '요청 %s' % r.get('no', ''), '요청 「%s」에 반영된 장표가 없음' % r['항목'], '장표 묶음 추천(Claude 가 요청 유형→장표 묶음 표에서 고름)')
    c.request_table = rows
    c.status[18] = '후보(대응표는 보고서 끝, 판정은 Claude 가 확정)'


def c19_rfp_keywords(c):
    if not c.a.rfp_keywords:
        c.status[19] = '입력 없음(--rfp-keywords)'
        return
    purpose = [s for s in c.slides if re.search(r'목적|범위|방향|목표|Why', s['title'] + s['message'])]
    ptxt = ' '.join(s['title'] + s['message'] + s['all_text'] for s in purpose)
    alltxt = ' '.join(s['all_text'] + s['title'] for s in c.slides)
    for kw in c.a.rfp_keywords:
        if kw in ptxt:
            continue
        where = '다른 장표에는 있음' if kw in alltxt else '제안서 어디에도 없음'
        c.add(19, CONFIRM, None, 'RFP 핵심어', 'RFP 핵심어 「%s」 가 목적·방향 장표에 없음(%s)' % (kw, where), '목적·방향 장표에 핵심어 배치')
    c.status[19] = '실행(후보)'


def c21_staff_clients(c):
    n = 0
    for s in c.slides:
        if not re.search(r'투입\s*인력|수행\s*조직|인력\s*(?:구성|현황|이력)|Profile', s['title'] + s['message'], re.I):
            continue
        for tb in s['tables']:
            hd = tb['rows'][0] if tb['rows'] else []
            idx = [i for i, h in enumerate(hd) if re.search(r'주요\s*고객|Clients?|수행\s*실적|주요\s*프로젝트', h, re.I)]
            for i in idx:
                cells = [r[i] for r in tb['rows'][1:] if len(r) > i and r[i]]
                if cells:
                    n += 1
                    c.add(21, CONFIRM, s['no'], '인력 표 「%s」 열' % hd[i], '주요 고객 목록 후보: %s' % ' / '.join(x.replace('\n', ', ')[:40] for x in cells[:4]),
                          '고객 업종과 맞는 실적을 앞으로(Claude 가 업종 관련성 판단)')
    c.status[21] = '후보 %d건' % n


def c23_meeting(c):
    if not c.a.meeting:
        c.status[23] = '입력 없음(--meeting)'
        return
    items = json.load(open(c.a.meeting, encoding='utf-8'))
    alltxt = ' '.join(s['all_text'] + ' ' + s['title'] for s in c.slides)
    table = []
    for it in items:
        kws = it.get('keywords') or re.findall(r'[가-힣A-Za-z0-9]{2,}', it['지시'])
        found = [kw for kw in kws if kw in alltxt]
        st = '반영 후보(키워드 %d/%d 확인)' % (len(found), len(kws)) if found else '미반영 후보'
        table.append((it, st))
        if not found:
            c.add(23, CONFIRM, None, '회의 지시 %s' % it.get('id', ''), '지시 「%s」 의 키워드가 장표에 없음' % it['지시'], '반영 확인표에 미반영 표시')
    c.meeting_table = table
    c.status[23] = '후보(반영 확인표는 보고서 끝)'


def c28_pages(c):
    r = c.public
    if not r or not (r.get('본문_쪽수') or r.get('요약본_쪽수') or r.get('발표_분')):
        c.status[28] = '꺼짐(발주처 규정 없음)'
        return
    body = [s for s, k in zip(c.slides, c.kinds) if k == 'body' and not _is_ref(s)]
    lim = r.get('요약본_쪽수') if r.get('이_파일은_요약본') else r.get('본문_쪽수')
    if lim and len(c.slides) > lim:
        c.add(28, CONFIRM, None, '쪽수', '전체 %d장이 상한 %d쪽을 넘음(본문 %d장)' % (len(c.slides), lim, len(body)), '넘는 장수만큼 줄일 후보([참고]·흐름 장표) 제시')
    if r.get('발표_분'):
        per = len(body) / float(r['발표_분'])
        if per > 2.5:
            c.add(28, CONFIRM, None, '발표 시간', '본문 %d장 / 발표 %d분 = 분당 %.1f장(빠름)' % (len(body), r['발표_분'], per), '발표 분량에 맞춰 압축 후보')
    c.status[28] = '실행(후보)'


RUNNERS = [c1_old_names, c2_placeholders, c3_company_figures, c4_example_labels, c5_anonymous, c6_sources, c7_title_prefix,
           c8_section_dividers, c9_layout_names, c10_weeks, c11_volume, c12_diagrams, c13_ref_split, c14_message,
           c15_assertive, c16_why, c17_endings, c18_requests, c19_rfp_keywords, c20_terms, c21_staff_clients,
           c22_price, c23_meeting, c24_anonymous, c25_track_record, c26_required, c27_format, c28_pages]


def render(c, path):
    L = {FIX: [], CONFIRM: [], REF: []}
    for f in c.f:
        L[f['level']].append(f)
    out = ['# 점검 보고서 · %s' % os.path.basename(c.d['file']), '',
           '장표 %d장 · 점검일 %s%s' % (c.d['slide_count'], c.today.isoformat(), (' · 새 고객 ' + c.client) if c.client else ''), '',
           '- 고칠 것 %d · 확인할 것 %d · 참고 %d' % (len(L[FIX]), len(L[CONFIRM]), len(L[REF])), '']
    for lvl, desc in ((FIX, '규칙으로 확정된 것 — 도구가 고치거나 고칠 안을 냄'), (CONFIRM, '후보 — 사람이(또는 Claude 가 문장을 읽고) 판단'), (REF, '알려 두는 것')):
        out += ['## %s (%d) — %s' % (lvl, len(L[lvl]), desc), '']
        if not L[lvl]:
            out += ['없음', '']
            continue
        out += ['| 항목 | 장 | 칸 | 내용 | 고치는 법 |', '|---|---|---|---|---|']
        for f in sorted(L[lvl], key=lambda x: (x['item'], x['slide'] or 0)):
            out.append('| %d %s | %s | %s | %s | %s |' % (f['item'], ITEMS[f['item']][0], ('#%d' % f['slide']) if f['slide'] else '전체',
                                                       f['where'].replace('|', '/'), f['msg'].replace('|', '/').replace('\n', ' '), f['suggest'].replace('|', '/')))
        out.append('')
    if getattr(c, 'request_table', None):
        out += ['## 요청 대응표 (후보)', '', '| # | 요청 항목 | 장표 | 판정 |', '|---|---|---|---|']
        for r, hit, v in c.request_table:
            out.append('| %s | %s | %s | %s |' % (r.get('no', ''), r['항목'], ', '.join('#%d' % s['no'] for s in hit[:8]) or '—', v))
        out.append('')
    if getattr(c, 'meeting_table', None):
        out += ['## 반영 확인표 (후보)', '', '| 지시 | 판정 |', '|---|---|']
        for it, st in c.meeting_table:
            out.append('| %s | %s |' % (it['지시'], st))
        out.append('')
    out += ['## 점검 현황 (28항목)', '', '| 번호 | 항목 | 처리 | 상태 |', '|---|---|---|---|']
    for n, (nm, mode, grp) in ITEMS.items():
        out.append('| %d | %s | %s | %s |' % (n, nm, mode, c.status.get(n, '실행 안 됨')))
    out.append('')
    open(path, 'w', encoding='utf-8').write('\n'.join(out))


def parse_weeks(v):
    m = re.fullmatch(r'(\d+)\s*[-~]\s*(\d+)', v)
    if not m:
        raise argparse.ArgumentTypeError('예: 8-12')
    return int(m.group(1)), int(m.group(2))


def main():
    ap = argparse.ArgumentParser(description='HCG 제안서 점검 28항목')
    ap.add_argument('pptx')
    ap.add_argument('--out', required=True)
    ap.add_argument('--client')
    ap.add_argument('--old-names')
    ap.add_argument('--figures')
    ap.add_argument('--public')
    ap.add_argument('--terms')
    ap.add_argument('--requests')
    ap.add_argument('--rfp-keywords', type=lambda v: [x.strip() for x in v.split(',') if x.strip()])
    ap.add_argument('--meeting')
    ap.add_argument('--pages', type=int)
    ap.add_argument('--rfp-weeks', type=parse_weeks)
    ap.add_argument('--json')
    a = ap.parse_args()
    # 따로 주지 않으면 스킬에 든 고객명 사전 · 회사 수치 표를 씀
    ref = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'references')
    if not a.old_names and os.path.exists(os.path.join(ref, 'client_names.txt')):
        a.old_names = os.path.join(ref, 'client_names.txt')
    if not a.figures and os.path.exists(os.path.join(ref, 'company_figures.json')):
        a.figures = os.path.join(ref, 'company_figures.json')
    d = extract(a.pptx)
    c = Ctx(d, a)
    for fn in RUNNERS:
        fn(c)
    render(c, a.out)
    if a.json:
        json.dump(c.f, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    n = Counter(f['level'] for f in c.f)
    print('고칠 것 %d · 확인할 것 %d · 참고 %d → %s' % (n[FIX], n[CONFIRM], n[REF], a.out))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
