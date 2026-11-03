// slotops.js 30장 시험: node test/slotops.test.mjs  (app/pane 에서)
// 읽기: app/store/out/pptx · meta.json (원본, 읽기만) / 쓰기: app/store/test3 · test/report_30.md
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';

// package.json 이 type:module 이라 .js 를 require 하면 ESM 취급 → UMD 를 module 객체를 넘겨 직접 실행
const HERE0 = path.dirname(fileURLToPath(import.meta.url));
const mod = { exports: {} };
new Function('module', 'exports', fs.readFileSync(path.join(HERE0, '../public/pane/slotops.js'), 'utf8'))(mod, mod.exports);
const S = mod.exports;
S.init({ DOMParser, XMLSerializer });

const HERE = path.dirname(fileURLToPath(import.meta.url));
const STORE = path.resolve(HERE, '../../store');
const OUT = path.join(STORE, 'out');
const TEST3 = path.join(STORE, 'test3');
fs.mkdirSync(TEST3, { recursive: true });

const meta = JSON.parse(fs.readFileSync(path.join(OUT, 'meta.json'), 'utf8'));
const parse = (x) => new DOMParser().parseFromString(x, 'application/xml');
const countIds = (xml) => parse(xml).getElementsByTagName('p:cNvPr').length;

async function load(id) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(OUT, 'pptx', id + '.pptx')));
  const sl = Object.keys(zip.files).find((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  const xml = await zip.file(sl).async('string');
  const opts = S.slideSize(await zip.file('ppt/presentation.xml').async('string'));
  return { zip, sl, xml, opts };
}

// ---- 장표: 1차 시험과 같은 30장(구조별로 골랐던 것)
const IDS = ["s00929","s01357","s05891","s12723","s15530","s00198","s03263","s06202","s06206","s07219","s03391","s13302","s03392","s01098","s04047","s10548","s08935","s23653","s01942","s03384","s09489","s01010","s08899","s04673","s11645","s03667","s04801","s05369","s07076","s08755"];
const byId = new Map(meta.map((m) => [m.id, m]));
const picks = IDS.map((id) => byId.get(id));

// ---- 검사
function checkDoc(xml, opts, before) {
  const errs = [], warns = [];
  let doc;
  try { doc = parse(xml); if (doc.getElementsByTagName('parsererror').length) throw new Error('parsererror'); }
  catch (e) { return { errs: ['다시 해석 실패: ' + e.message], warns }; }
  if (!/xmlns:p=/.test(xml) || !/xmlns:a=/.test(xml)) errs.push('네임스페이스 선언 사라짐');
  const ids = [...doc.getElementsByTagName('p:cNvPr')].map((c) => c.getAttribute('id'));
  const dupOf = (x) => { const a = [...parse(x).getElementsByTagName('p:cNvPr')].map((c) => c.getAttribute('id')); return a.length - new Set(a).size; };
  if (dupOf(xml) > dupOf(before)) errs.push('도형 번호(id) 중복 새로 생김');
  for (const g of [...doc.getElementsByTagName('p:grpSp')]) {
    const x = [...g.childNodes].find((n) => n.localName === 'grpSpPr');
    const xf = x && [...x.childNodes].find((n) => n.localName === 'xfrm');
    if (!xf) continue;
    const v = {};
    for (const n of ['off', 'ext', 'chOff', 'chExt']) {
      const e = [...xf.childNodes].find((c) => c.localName === n);
      for (const a of ['x', 'y', 'cx', 'cy']) if (e && e.hasAttribute(a)) { v[n + a] = Number(e.getAttribute(a)); if (!Number.isFinite(v[n + a])) errs.push('묶음 좌표 비정상'); }
    }
    if (v.extcx < 0 || v.chExtcx < 0 || v.extcx === 0 && v.chExtcx > 0) errs.push('묶음 폭 비정상');
  }
  const after = S.slotMap(xml, opts);
  const tol = 0.01 * opts.slideW;
  const out = (s) => s.w > 0 && (s.x < -tol || s.y < -tol || s.x + s.w > opts.slideW + tol || s.y + s.h > opts.slideH + tol);
  const outBefore = new Set(S.slotMap(before, opts).filter(out).map((s) => s.id));
  const newOut = after.filter((s) => out(s) && !outBefore.has(s.id));
  if (newOut.length) errs.push('장표 밖으로 나간 도형 ' + newOut.length + '개(#' + newOut.map((s) => s.no).slice(0, 4).join(',') + ')');
  return { errs, warns, after };
}
function colOverlap(xml, opts, n) {
  const cs = S.columns(xml, opts);
  if (cs.length !== n) return '열 재감지 ' + cs.length + '개(기대 ' + n + ')';
  for (let i = 0; i + 1 < cs.length; i++) {
    const ov = cs[i].x1 - cs[i + 1].x0, nar = Math.min(cs[i].x1 - cs[i].x0, cs[i + 1].x1 - cs[i + 1].x0);
    if (ov > 0.05 * nar) return '열 ' + (i + 1) + '·' + (i + 2) + ' 가로 겹침 ' + Math.round(ov / nar * 100) + '%';
  }
  return null;
}
function memberCount(xml, ids) {
  const doc = parse(xml); let n = 0;
  const set = new Set(ids.map(String));
  const seen = new Set();
  for (const c of [...doc.getElementsByTagName('p:cNvPr')]) {
    if (!set.has(c.getAttribute('id')) || seen.has(c.getAttribute('id'))) continue;
    seen.add(c.getAttribute('id'));
    n += c.parentNode.parentNode.getElementsByTagName('p:cNvPr').length; // 도형 자신 + 안쪽 도형
  }
  return n;
}

async function save(zip, sl, xml, id, op) {
  zip.file(sl, xml);
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(path.join(TEST3, id + '_' + op + '.pptx'), buf);
}

// ---- 본문
const rows = [];
let nOk = 0, nWarn = 0, nFail = 0, nCol3 = 0, nTbl = 0;
const t0 = Date.now();

for (const m of picks) {
  const row = { id: m.id, s: m.s + ' ' + m.s_name, slots: 0, cols: 0, ops: [], res: 'OK', warns: [] };
  rows.push(row);
  let fail = false;
  try {
    const { zip, sl, xml, opts } = await load(m.id);
    const map = S.slotMap(xml, opts);
    const cs = S.columns(xml, opts);
    const ov = S.overlaps(map);
    row.slots = map.length; row.cols = cs.length;
    row.over = map.filter((s) => s.overflow).length; row.ovl = ov.length;
    if (map.some((s) => !Number.isFinite(s.x + s.y + s.w + s.h))) { fail = true; row.warns.push('좌표 NaN'); }
    const n = cs.length;
    const info = S.columnsInfo(xml, opts);
    if (!n && info.reason && !/2열 미만/.test(info.reason)) {
      const r = S.setColumns(xml, 3, opts);
      const okRej = r.xml === xml && r.warnings.length > 0;
      row.ops.push('열 판정 거절: ' + info.reason);
      if (!okRej) { fail = true; row.warns.push('오류 거절했는데 바뀜/경고 없음'); }
    }
    if (n >= 3) {
      nCol3++;
      const tries = [['열-1', n - 1]];
      if (n + 1 <= 5) tries.push(['열+1', n + 1]);
      for (const [lab, to] of tries) {
        const r = S.setColumns(xml, to, opts);
        const op = lab + '(' + n + '→' + to + ')';
        const before = countIds(xml);
        if (r.xml === xml) { row.ops.push(op + ' 실패'); row.warns.push(op + ': ' + r.warnings.join('/')); fail = true; continue; }
        const ck = checkDoc(r.xml, opts, xml);
        const after = countIds(r.xml);
        const mc = memberCount(xml, cs[n - 1].members);
        if (to < n && before - after < mc) ck.errs.push('도형 수 기대보다 적게 줄음(' + before + '→' + after + ', 열 ' + mc + ')');
        if (to > n && (after - before !== r.stats.added || after - before < mc)) ck.errs.push('도형 수 기대와 다름(' + before + '→' + after + ', 열 ' + mc + ')');
        const co = colOverlap(r.xml, opts, to);
        if (co) ck.warns.push(co);
        row.ops.push(op + ' 도형 ' + before + '→' + after);
        ck.errs.forEach((e) => { row.warns.push('오류 ' + op + ': ' + e); fail = true; });
        ck.warns.concat(r.warnings).forEach((w) => row.warns.push(op + ': ' + w));
        if (!ck.errs.length) await save(zip, sl, r.xml, m.id, lab.replace(/[^0-9a-z+\-가-힣]/g, '') + to);
        else await save(zip, sl, r.xml, m.id, 'BAD_' + lab.replace(/[^0-9a-z+\-가-힣]/g, '') + to);
      }
    }
    const tb = map.find((s) => s.kind === 'table' && s.rows >= 3);
    if (tb) {
      nTbl++;
      let done = false;
      const cand = [];
      for (let r = Math.floor(tb.rows / 2); r < tb.rows; r++) cand.push(r);
      for (let r = Math.floor(tb.rows / 2) - 1; r >= 1; r--) cand.push(r);
      const rejected = [];
      for (const r of cand) {
        const d = S.deleteTableRow(xml, tb.no, r, opts);
        if (d.xml === xml) { rejected.push(r + ':' + d.warnings.join('')); continue; }
        const doc = parse(d.xml), od = parse(xml);
        const f = (x) => [...x.getElementsByTagName('a:tbl')].map((t) => {
          const trs = [...t.childNodes].filter((c) => c.localName === 'tr');
          return { n: trs.length, h: trs.reduce((a, b) => a + Number(b.getAttribute('h')), 0) };
        });
        const tn = tb.tableIdx ?? 0;
        const a = f(od), b = f(doc);
        const idx = a.findIndex((v, i) => v.n !== b[i].n);
        const errs = [];
        if (idx < 0 || a[idx].n - 1 !== b[idx].n) errs.push('행 수 −1 아님');
        else if (Math.abs(a[idx].h - b[idx].h) > 0) errs.push('행 높이 합 달라짐 ' + a[idx].h + '→' + b[idx].h);
        const ck = checkDoc(d.xml, opts, xml);
        errs.push(...ck.errs);
        row.ops.push('행삭제(표#' + tb.no + ' 행' + r + ', ' + tb.rows + '→' + (tb.rows - 1) + '행)' + (rejected.length ? ' [앞서 거절 ' + rejected.length + ']' : ''));
        if (rejected.length) row.warns.push('병합 거절: ' + rejected.slice(0, 2).join(' | '));
        errs.forEach((e) => { row.warns.push('오류 행삭제: ' + e); fail = true; });
        await save(zip, sl, d.xml, m.id, (errs.length ? 'BAD_' : '') + '행삭제' + r);
        done = true;
        break;
      }
      if (!done) { row.ops.push('행삭제 전부 거절'); row.warns.push('병합 때문에 지울 행 없음: ' + rejected.slice(0, 2).join(' | ')); }
    }
  } catch (e) {
    fail = true; row.warns.push('예외: ' + e.message.slice(0, 160));
  }
  row.res = fail ? '실패' : (row.warns.length ? '경고' : '성공');
  if (fail) nFail++; else if (row.warns.length) nWarn++; else nOk++;
}

const lines = [
  '# slotops.js 30장 시험 보고', '',
  '- 시험 ' + rows.length + '장 · 성공 ' + nOk + ' · 경고 ' + nWarn + ' · 실패 ' + nFail + ' · 열 3개 이상 ' + nCol3 + '장 · 표(3행 이상) ' + nTbl + '장 · ' + ((Date.now() - t0) / 1000).toFixed(1) + '초',
  '- 바꾼 장표 저장: app/store/test3/<id>_<조작>.pptx (BAD_ 로 시작하면 검사 실패본)',
  '- 넘침·겹침은 어림값(글자 폭 × 글꼴 크기 기준)', '',
  '| id | 구조 | 칸 | 열 | 넘침 | 겹침 | 한 조작 | 결과 | 경고 |', '|---|---|---|---|---|---|---|---|---|'
];
for (const r of rows) {
  lines.push('| ' + [r.id, r.s, r.slots, r.cols, r.over ?? '', r.ovl ?? '', r.ops.join('; ') || '(조작 없음)', r.res, r.warns.join(' / ') || '-'].join(' | ') + ' |');
}
fs.writeFileSync(path.join(HERE, 'report_30.md'), lines.join('\n') + '\n', 'utf8');
console.log(lines.slice(2, 4).join('\n'));
console.log(rows.map((r) => r.id + ' ' + r.s + ' 칸' + r.slots + ' 열' + r.cols + ' ' + r.res + (r.warns.length ? ' :: ' + r.warns[0] : '')).join('\n'));
process.exit(nFail ? 1 : 0);
