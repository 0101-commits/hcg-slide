// checkops.js 대조 시험: node test/checkops.test.mjs  (app/pane 에서)
// 같은 pptx 에 파이썬 판(check.py)과 JS 판을 돌려 항목별 건수·(항목,수준,장,칸) 목록이 같은지 비교.
// 읽기: work/converted · work/spike_struct 의 pptx (원본, 읽기만) / 쓰기: 임시 폴더 · test/report_check.md
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const mod = { exports: {} };
new Function('module', 'exports', fs.readFileSync(path.join(HERE, '../public/pane/checkops.js'), 'utf8'))(mod, mod.exports);
const C = mod.exports;
C.init({ DOMParser });

const ROOT = path.resolve(HERE, '../../..');
const PY = path.join(ROOT, '.venv/Scripts/python.exe');
const SK = path.join(ROOT, 'app/skills/hcg-check');
const CHECK = path.join(SK, 'scripts/check.py');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'checkops-'));
const CLIENT = '한빛소재';
const LV = { '고칠 것': 'fix', '확인할 것': 'confirm', '참고': 'ref' };

const oldNames = fs.readFileSync(path.join(SK, 'references/client_names.txt'), 'utf8').split(/\r?\n/).map((x) => x.trim()).filter((x) => x && !x.startsWith('#'));
const figures = JSON.parse(fs.readFileSync(path.join(SK, 'references/company_figures.json'), 'utf8'));

// 옵션을 모두 켠 두 번째 실행용 입력(파이썬·JS 에 같은 값)
const EXTRA = {
  terms: { '성과': '성과급', '인센티브': '성과급', '고객': '고객사' },
  requests: [{ no: 1, 항목: '성과평가 체계 개선' }, { no: 2, 항목: '교육 체계', keywords: ['교육'] }, { no: 3, 항목: '별도 가격', 본문밖: true }, { no: 4, 항목: '조직문화 진단 설문' }],
  meeting: [{ id: 'm1', 지시: '일정 압축 반영', keywords: ['일정', '압축'] }, { id: 'm2', 지시: '해외법인 사례 추가' }],
  rfp: ['성과', '역량', '존재하지않는말'],
  pages: 30,
  weeks: [8, 12],
  pub: { 무기명: true, 금지어: ['한빛'], 금지_인명: ['홍길동'], 최근_N년: 3, 기준연도: 2026, 최소_실적건수: 5, 지정_목차: ['사업 이해', '수행 전략', '사업 관리'],
    필수_장표: ['업무 보안', '사후 관리'], 지정_레이아웃: ['표지', '본문'], 인력표_열: ['성명', '직급', '경력', '주요 고객'], 본문_쪽수: 30, 요약본_쪽수: 20, 발표_분: 10 }
};
const J = (n, o) => { const p = path.join(TMP, n); fs.writeFileSync(p, JSON.stringify(o), 'utf8'); return p; };
const EX = { terms: J('terms.json', EXTRA.terms), requests: J('requests.json', EXTRA.requests), meeting: J('meeting.json', EXTRA.meeting), pub: J('pub.json', EXTRA.pub) };

async function loadFiles(p) {
  const zip = await JSZip.loadAsync(fs.readFileSync(p));
  const files = {};
  for (const n of Object.keys(zip.files)) if (/\.(xml|rels)$/.test(n) && !zip.files[n].dir) files[n] = await zip.file(n).async('string');
  return files;
}

function runPy(pptx, tag, extra) {
  const out = path.join(TMP, tag + '.json');
  const r = spawnSync(PY, ['-X', 'utf8', CHECK, pptx, '--client', CLIENT, '--json', out, '--out', path.join(TMP, tag + '.md')].concat(extra), { encoding: 'utf8', env: { ...process.env, PYTHONUTF8: '1' } });
  if (r.status !== 0) throw new Error('파이썬 실패: ' + pptx + '\n' + r.stderr);
  return JSON.parse(fs.readFileSync(out, 'utf8')).map((f) => ({ item: f.item, level: LV[f.level], slide: f.slide || null, shape: f.where }));
}

function runJs(files, extra) {
  const o = { client: CLIENT, oldNames, figures, publicRules: null };
  if (extra) Object.assign(o, { terms: EXTRA.terms, requests: EXTRA.requests, meeting: EXTRA.meeting, rfpKeywords: EXTRA.rfp, pages: EXTRA.pages, rfpWeeks: EXTRA.weeks, publicRules: EXTRA.pub });
  return C.run(files, o).findings.map((f) => ({ item: f.item, level: f.level, slide: f.slide, shape: f.shape }));
}

const keyOf = (f) => [f.item, f.level, f.slide || '', f.shape || ''].join('|');
function tally(list, fn) { const m = new Map(); list.forEach((f) => { const k = fn(f); m.set(k, (m.get(k) || 0) + 1); }); return m; }

function compare(py, js) {
  const byItem = (f) => f.item + '/' + f.level;
  const a = tally(py, byItem), b = tally(js, byItem);
  const keys = Array.from(new Set([...a.keys(), ...b.keys()])).sort((x, y) => parseInt(x) - parseInt(y) || x.localeCompare(y));
  const rows = keys.map((k) => ({ k, py: a.get(k) || 0, js: b.get(k) || 0 }));
  const ka = tally(py, keyOf), kb = tally(js, keyOf), diffs = [];
  new Set([...ka.keys(), ...kb.keys()]).forEach((k) => { if ((ka.get(k) || 0) !== (kb.get(k) || 0)) diffs.push(k + '  py=' + (ka.get(k) || 0) + ' js=' + (kb.get(k) || 0)); });
  return { rows, diffs, same: rows.every((r) => r.py === r.js) };
}

// 시험 파일: 여러 장짜리(8장 이상) 위주
function candidates() {
  const list = [];
  const conv = path.join(ROOT, 'work/converted');
  for (const d of fs.readdirSync(conv)) {
    const dir = path.join(conv, d);
    if (!fs.statSync(dir).isDirectory()) continue;
    for (const f of fs.readdirSync(dir)) if (f.endsWith('.pptx')) list.push(path.join(dir, f));
  }
  const sp = path.join(ROOT, 'work/spike_struct');
  for (const d of fs.readdirSync(sp)) {
    const p = path.join(sp, d, 'orig.pptx');
    if (/^s\d/.test(d) && fs.existsSync(p)) list.push(p);
  }
  return list;
}

const WANT_CONV = 14;
const picked = [];
let nConv = 0;
for (const p of candidates()) {
  const isSp = p.includes('spike_struct');
  if (!isSp && nConv >= WANT_CONV) continue;
  let files;
  try { files = await loadFiles(p); } catch (e) { continue; }
  let n = 0;
  try { n = C.extract(files).length; } catch (e) { continue; }
  if (!isSp && n < 8) continue;
  picked.push({ p, files, n, isSp });
  if (!isSp) nConv++;
}

const out = ['# checkops.js 대조 보고(파이썬 판 check.py vs JS 판)', '',
  '- 고객 ' + CLIENT + ' · 공식 수치 2026.04 · 이전 고객명 사전 ' + oldNames.length + '건 · 시험 파일 ' + picked.length + '개',
  '- 실행 A = 기본 입력(고객명·수치 표만). 실행 B = 용어 사전·요청 항목·회의 지시·RFP 핵심어·분량·RFP 주수·공공 규정(무기명 등)까지 모두 켬',
  '- 대조 기준: 항목/수준별 건수, 그리고 (항목·수준·장·칸) 목록 일치', ''];
let okAll = 0, total = 0;
const summary = [];
for (const it of picked.slice(0, process.env.CHECK_LIMIT ? Number(process.env.CHECK_LIMIT) : picked.length)) {   // CHECK_LIMIT: 빠른 시험용(보고서는 전체로 다시 만들 것)
  const label = path.relative(ROOT, it.p).replace(/\\/g, '/');
  const sec = ['## ' + label + ' (' + it.n + '장)', ''];
  for (const [tag, extra] of [['A', false], ['B', true]]) {
    const pyArgs = extra ? ['--terms', EX.terms, '--requests', EX.requests, '--meeting', EX.meeting, '--public', EX.pub, '--pages', '30', '--rfp-weeks', '8-12', '--rfp-keywords', EXTRA.rfp.join(',')] : [];
    const py = runPy(it.p, 'o', pyArgs), js = runJs(it.files, extra), r = compare(py, js);
    total++; if (r.same && !r.diffs.length) okAll++;
    sec.push('실행 ' + tag + ': 파이썬 ' + py.length + '건 · JS ' + js.length + '건 → ' + (r.same && !r.diffs.length ? '전부 일치' : (r.same ? '건수 일치, 목록 차이 ' + r.diffs.length : '다름')), '');
    sec.push('| 항목/수준 | 파이썬 | JS | 일치 |', '|---|---|---|---|');
    r.rows.forEach((x) => sec.push('| ' + x.k + ' | ' + x.py + ' | ' + x.js + ' | ' + (x.py === x.js ? 'O' : 'X') + ' |'));
    if (!r.rows.length) sec.push('| (없음) | 0 | 0 | O |');
    if (r.diffs.length) { sec.push('', '목록 차이(앞 15):'); r.diffs.slice(0, 15).forEach((d) => sec.push('- ' + d)); }
    sec.push('');
    summary.push(label + ' ' + tag + ': ' + (r.same && !r.diffs.length ? 'OK' : 'DIFF'));
  }
  out.push(...sec);
}
// ---- 합성 시험: 실제 파일에 잘 안 나오는 규칙(고객명 묶기·기준월 있는/없는 수치·습니다체 인사말/인용/전체)을 직접 넣어 대조
{
  const base = picked.find((x) => x.n >= 8);
  const zip = await JSZip.loadAsync(fs.readFileSync(base.p));
  const slideFiles = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
  const names6 = oldNames.filter((n) => n !== CLIENT && /^[가-힣]{3,5}$/.test(n)).slice(0, 6);
  const sp = (id, y, lines) => '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="합성 ' + id + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="' + y + '"/><a:ext cx="9000000" cy="900000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>' +
    lines.map((t) => '<a:p><a:r><a:t>' + t + '</a:t></a:r></a:p>').join('') + '</p:txBody></p:sp>';
  const polite = ['이번 제안은 귀사의 성과 관리 체계를 개선하기 위한 것입니다.', '평가 결과는 보상 의사결정에 직접 연결해 운영합니다.', '진행 과정은 매주 보고서로 공유드립니다.'];
  const extra = [
    sp(9001, 3000000, names6.map((n) => n + ' 프로젝트 수행')),                       // 사전 이름 6개 → 한 건으로 묶임
    sp(9002, 3200000, ['고객 950사 (2026.01 기준) 보유한 서비스 현황입니다', '고객 820사 보유한 서비스 현황을 소개함', '고객 700개사 (2025년 1월 현재) 보유 현황임']),
    sp(9003, 3400000, polite.concat(['끝까지 경청해 주셔서 진심으로 감사합니다.', '고객은 "빠른 대응과 안정적인 운영이 필요합니다']))
  ].join('');
  const extra2 = sp(9004, 3000000, names6.slice(0, 4).map((n) => n + ' 프로젝트 수행'));    // 4개 → 이름마다 고칠 것
  const tmpP = path.join(TMP, 'synthetic.pptx');
  for (let i = 1; i < slideFiles.length; i++) {
    const n = slideFiles[i];
    zip.file(n, (await zip.file(n).async('string')).replace('</p:spTree>', (i % 2 ? extra : extra2) + '</p:spTree>'));
  }
  fs.writeFileSync(tmpP, await zip.generateAsync({ type: 'nodebuffer' }));
  const sfiles = await loadFiles(tmpP);
  const py = runPy(tmpP, 'syn', []), jsr = C.run(sfiles, { client: CLIENT, oldNames, figures, publicRules: null }), r = compare(py, jsr.findings.map((f) => ({ item: f.item, level: f.level, slide: f.slide, shape: f.shape })));
  total++; if (r.same && !r.diffs.length) okAll++;
  const fd = jsr.findings;
  const n1c = fd.filter((f) => f.item === 1 && f.level === 'confirm' && /^고객명 \d+개/.test(f.text)).length;
  const n1f = fd.filter((f) => f.item === 1 && f.level === 'fix' && f.from).length;
  const n3f = fd.filter((f) => f.item === 3 && f.level === 'fix' && f.from === '950').length, n3c = fd.filter((f) => f.item === 3 && f.level === 'confirm' && f.from === '820').length;
  const n17c = fd.filter((f) => f.item === 17 && f.level === 'confirm').length, n17g = fd.filter((f) => f.item === 17 && /감사합니다|필요합니다/.test(f.text)).length;
  // 창고 태그: 같은 장에 이름 6개가 있어도 태그 이름은 고칠 것으로 따로
  const tagRun = C.run(sfiles, { client: CLIENT, oldNames, figures, slideTags: { 2: { HCG_OLD: names6[0] } } }).findings.filter((f) => f.item === 1 && f.slide === 2);
  const tagFix = tagRun.filter((f) => f.level === 'fix' && f.from).length;
  out.push('## 합성 시험(고객명 묶기·기준월·습니다체)', '', '실행: 파이썬 ' + py.length + '건 · JS ' + jsr.findings.length + '건 → ' + (r.same && !r.diffs.length ? '전부 일치' : '다름'), '',
    '- 고객명 묶음(확인할 것 「고객명 N개」) ' + n1c + '건 · 4개 이하 장표의 이름별 고칠 것 ' + n1f + '건 · 태그 이름이 든 장표의 고칠 것 ' + tagFix + '건',
    '- 기준월 붙은 「고객 950사」 고칠 것 ' + n3f + '건 · 기준월 없는 「고객 820사」 확인할 것 ' + n3c + '건',
    '- 습니다체 확인할 것 ' + n17c + '건(인사말·인용 포함 문장이 잡힌 건수 ' + n17g + ')', '');
  if (!(n1c > 0 && n1f > 0 && tagFix === 1 && n3f > 0 && n3c > 0 && n17c > 0 && n17g === 0)) { console.log('합성 시험 기대값 어긋남', { n1c, n1f, tagFix, n3f, n3c, n17c, n17g }); okAll--; }
  summary.push('합성 시험: ' + (r.same && !r.diffs.length ? 'OK' : 'DIFF') + ' ' + JSON.stringify({ n1c, n1f, tagFix, n3f, n3c, n17c, n17g }));
}
out.splice(5, 0, '- 결과: ' + okAll + ' / ' + total + ' 실행이 완전 일치');
fs.writeFileSync(path.join(HERE, 'report_check.md'), out.join('\n'), 'utf8');
console.log(summary.join('\n'));
console.log('완전 일치 ' + okAll + ' / ' + total);
process.exitCode = okAll === total ? 0 : 1;
