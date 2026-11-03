"""고객명 사전 만들기 (관리자가 다시 돌림).

과거 제안 폴더 이름(projects.json)과 참고 제안서 파일 이름(ref/files.json)에서 고객명을 뽑아
hcg-check/references/client_names.txt (한 줄에 하나, 중복 제거, 일반 낱말 뺌)를 만든다.
이 목록은 hcg-check 의 「이전 고객명 잔재」 검사와 hcg-restructure/residue.py 가 쓴다.

사용:
  python make_client_names.py [--projects ...\\projects.json] [--files ...\\files.json] [--out ...\\client_names.txt]
  (인자를 안 주면 아래 기본 경로를 씀)
새 제안·참고 파일이 늘면 다시 돌리고, 결과를 눈으로 한 번 훑어 일반 낱말이 섞였으면 STOP 에 보탠다.
"""
import argparse
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEF_PROJECTS = os.path.join(BASE, 'work', 'projects.json')
DEF_FILES = os.path.join(BASE, 'work', 'ref', 'files.json')
DEF_OUT = os.path.join(BASE, 'app', 'skills', 'hcg-check', 'references', 'client_names.txt')

# 파일 이름 앞에 붙는 꼬리표(고객명이 아님)
TAGS = ('HCG', 'HCC', 'IBM', '네모', '별첨1', '별첨')
# 고객명이 아닌 일반 낱말·자사 이름·자료 종류 (검사 때 오탐을 막으려고 뺌)
STOP = {
    '관리방안', '한계인력', '핵심&저성과', '핵심', '회사소개', '회사', '인사제도', '프로젝트', '참고', '직무관리',
    '시스템', 'HR제도', 'HR컨설팅', 'HR', '임원', '리더십', '리더십&조직문화', '조직', '인사', '평가', '보상',
    '휴넬', '휴넬-퍼플', 'hunel', 'talenx', 'JaDE', 'DNapp', 'Performance', 'SPO', 'PM', 'PPT', 'Time', '제안서',
    '가격', '견적', '요약본', '전체본', '발표본', '약식제안서', '일하는', '신', '전반부', '산출물', '급여',
    '매일', 'pm프로필', '회사소개서', 'Project', 'Company', 'Reference', 'Template', '템플릿', '표준', '그래픽', 'Graphic',
}
# 폴더·파일 이름에서 못 뽑는 줄임말·영문 표기 (파일 이름에 근거가 있는 것만)
EXTRA = ['mySUNI', 'W concept']


def clean(raw):
    s = raw.replace('\\', '/').split('/')[-1]
    s = re.sub(r'\.(pptx?|docx?|xlsx?)$', '', s, flags=re.I)
    s = s.lstrip('@')
    s = re.sub(r'^\d{6}_', '', s)               # 150826_하나대투증권_...
    s = re.sub(r'^r_템플릿_', '', s)
    changed = True
    while changed:                                # [HCG] · HCG_ · HCG  꼬리표 반복 제거
        changed = False
        m = re.match(r'^\[([^\]]+)\]\s*(.*)$', s)
        if m and (m.group(1).strip() in TAGS or re.fullmatch(r'별첨\d*', m.group(1).strip())):
            s, changed = m.group(2), True
            continue
        m = re.match(r'^\[([^\]]+)\]\s*(.*)$', s)
        if m:                                     # [녹십자홀딩스] 성과관리 ... → 괄호 안이 고객명
            s, changed = m.group(1) + '_' + m.group(2), True
            continue
        m = re.match(r'^(?:' + '|'.join(TAGS) + r')(?:[_\s]+|(?=[가-힣]))(.*)$', s)
        if m:
            s, changed = m.group(1), True
    return s.strip()


def from_file_name(raw):
    s = clean(raw)
    if s.startswith('(') or not s:
        return None
    seg = re.split(r'[_]', s)[0].strip()
    seg = seg.replace('㈜', '').replace('(주)', '').strip()
    if ' ' in seg and len(seg) > 12:
        seg = seg.split()[0]
    elif ' ' in seg:
        first = seg.split()[0]
        # 「kt estate」「DB Inc」처럼 짧은 영문 두 단어는 통째로, 그 밖은 첫 낱말만
        if not re.fullmatch(r'[A-Za-z]+( [A-Za-z]+)?', seg):
            seg = first
    seg = re.sub(r'(?<=[가-힣])의$', '', seg)    # 「기아의 가치와」 → 기아
    return seg


def from_project_name(name):
    s = re.sub(r'^\d{4}\.\d{2}\.\d{2}_', '', name)
    s = s.split('_')[0]
    s = re.sub(r'\([^)]*\)', '', s)
    s = re.sub(r'\s*(장기계약|프로젝트 정리)$', '', s).strip()
    return s


def keep(n):
    if not n or n in STOP or n.lower() in {x.lower() for x in STOP}:
        return False
    if len(n) > 10:                               # 과제 이름이 통째로 들어온 것
        return False
    ascii_only = re.fullmatch(r'[A-Za-z0-9 .&\-]+', n) is not None
    if ascii_only and len(n.replace(' ', '')) < 3:   # CR · DB · AP 같은 짧은 영문은 오탐이 많아 뺌
        return False
    if not ascii_only and len(re.sub(r'[^가-힣]', '', n)) < 2:
        return False
    if re.search(r'\d{4,}|\d\.', n):
        return False
    return True


def main():
    ap = argparse.ArgumentParser(description='고객명 사전(client_names.txt) 만들기')
    ap.add_argument('--projects', default=DEF_PROJECTS)
    ap.add_argument('--files', default=DEF_FILES)
    ap.add_argument('--out', default=DEF_OUT)
    a = ap.parse_args()
    cands = []
    for p in json.load(open(a.projects, encoding='utf-8')):
        cands.append(from_project_name(p['name']))
    for f in json.load(open(a.files, encoding='utf-8')):
        cands.append(from_file_name(f['rel']))
    cands += EXTRA
    names, seen = [], set()
    for n in cands:
        if not n:
            continue
        n = n.strip()
        forms = [n]
        if ' ' in n and re.search(r'[가-힣]', n):    # 「KB 라이프생명」 → 「KB라이프생명」 도 함께
            forms.append(n.replace(' ', ''))
        # 「일진그룹」 → 「일진」 도 함께(본문에는 그룹을 뺀 이름으로 자주 나옴)
        base = re.sub(r'(그룹|홀딩스)$', '', n)
        if base != n and re.search(r'[가-힣]{2,}', base):
            forms.append(base)
        for x in forms:
            if keep(x) and x.lower() not in seen:
                seen.add(x.lower())
                names.append(x)
    names.sort(key=lambda x: (not re.search(r'[가-힣]', x), x))
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with open(a.out, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write('# HCG 과거 제안·참고 파일에서 뽑은 고객명 (한 줄에 하나). 관리자가 admin/make_client_names.py 로 다시 만듦.\n')
        fh.write('\n'.join(names) + '\n')
    print('고객명 %d개 → %s' % (len(names), a.out))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
