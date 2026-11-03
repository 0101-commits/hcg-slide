// 파이썬 판(slotmap.py)과 칸 수 · 열 수 대조: node test/parity.mjs  (사전: py -X utf8 slotmap.py 로 만든 slots.json 이 work 폴더에 있어야 함)
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const mod = { exports: {} };
new Function('module', 'exports', fs.readFileSync(path.join(HERE, '../public/pane/slotops.js'), 'utf8'))(mod, mod.exports);
const S = mod.exports;
S.init({ DOMParser, XMLSerializer });
const ROOT = path.resolve(HERE, '../../..');
const PY = path.join(ROOT, '.venv/Scripts/python.exe');
const SM = path.join(ROOT, 'app/skills/hcg-restructure/scripts/slotmap.py');
const TMP = process.env.TEMP || '.';
const out = [];
for (const d of ['s1_card', 's2_asis', 's3_option', 's4_gantt', 's5_org']) {
  const f = path.join(ROOT, 'work/spike_struct', d, 'orig.pptx');
  const js = path.join(TMP, 'parity_' + d + '.json');
  execFileSync(PY, ['-X', 'utf8', SM, f, '--out', js], { stdio: 'pipe' });
  const py = JSON.parse(fs.readFileSync(js, 'utf8')).slides[0];
  const zip = await JSZip.loadAsync(fs.readFileSync(f));
  const sl = Object.keys(zip.files).find((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  const xml = await zip.file(sl).async('string');
  const o = S.slideSize(await zip.file('ppt/presentation.xml').async('string'));
  const m = S.slotMap(xml, o), cs = S.columns(xml, o);
  const pyCols = (py.columns || []).length;
  const kinds = (a) => a.reduce((r, k) => ((r[k] = (r[k] || 0) + 1), r), {});
  const pk = kinds(py.slots.map((s) => s.kind)), jk = kinds(m.map((s) => (s.kind === 'pic' ? 'picture' : s.kind)));
  const same = m.length === py.slots.length && cs.length === pyCols;
  out.push(`${d}: 칸 py ${py.slots.length} / js ${m.length} · 열 py ${pyCols} / js ${cs.length} → ${same ? '같음' : '다름'}` + (same ? '' : ` | py ${JSON.stringify(pk)} js ${JSON.stringify(jk)}`));
}
console.log(out.join('\n'));
