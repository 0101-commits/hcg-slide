/* HCG 제안서 점검 엔진 — PPTX 를 푼 파일 목록(문자열 맵)만 다룸(JSZip · Office.js 몰라도 됨).
 * 파이썬 판(hcg-check/scripts/check.py · extract_text.py)의 28항목 규칙을 그대로 옮김.
 * 브라우저: window.HCGCheck · Node: module.exports. DOMParser/XMLSerializer 는 init 으로 넣을 수 있음.
 *
 *   HCGCheck.extract(files)        → [{no, layout, title, shapes:[{name,text,kind,x,y,w,h}], notes, hidden}]
 *   HCGCheck.run(files, opts)      → {findings:[{item,name,level,slide,shape,text,how,from?,to?}], status:{번호:문구}}
 *   opts = {client, oldNames[], figures{}, publicRules{}|null, slideTags{장:{HCG_OLD:"a|b"}},
 *           terms{}, requests[], meeting[], rfpKeywords[], pages, rfpWeeks:[lo,hi], today:Date}
 *   level: "fix"(고칠 것) | "confirm"(확인할 것) | "ref"(참고)
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HCGCheck = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DP = null;
  function init(o) {
    o = o || {};
    if (o.DOMParser) DP = o.DOMParser;
  }

  // ------------------------------------------------------------ XML 도우미
  function parse(xml) {
    var P = DP || (typeof DOMParser !== 'undefined' ? DOMParser : null);
    if (!P) throw new Error('DOMParser 가 없음: HCGCheck.init({DOMParser, XMLSerializer}) 를 먼저 부를 것');
    var doc = new P().parseFromString(xml, 'application/xml');
    var err = doc.getElementsByTagName('parsererror');
    if (err && err.length) throw new Error('XML 해석 실패: ' + err[0].textContent.slice(0, 120));
    if (!doc.documentElement) throw new Error('XML 해석 실패: 빈 문서');
    return doc;
  }
  function ln(el) { return el.localName || String(el.nodeName).replace(/^.*:/, ''); }
  function kids(el, name) {
    var out = [], c = el && el.childNodes;
    if (!c) return out;
    for (var i = 0; i < c.length; i++) if (c[i].nodeType === 1 && (!name || ln(c[i]) === name)) out.push(c[i]);
    return out;
  }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function descs(el, name, out) {
    out = out || [];
    var c = el.childNodes;
    for (var i = 0; i < c.length; i++) {
      if (c[i].nodeType !== 1) continue;
      if (ln(c[i]) === name) out.push(c[i]);
      descs(c[i], name, out);
    }
    return out;
  }
  function attr(el, name) {
    if (!el) return null;
    var v = el.getAttribute(name);
    if (v !== null && v !== '') return v;
    var a = el.attributes;
    for (var i = 0; a && i < a.length; i++) if (a[i].name === name || a[i].name.replace(/^.*:/, '') === name) return a[i].value;
    return v === '' ? '' : null;
  }
  function txt(el) { return el ? (el.textContent || '') : ''; }
  function hasFile(files, p) { return Object.prototype.hasOwnProperty.call(files, p); }

  function dirOf(p) { return p.indexOf('/') < 0 ? '' : p.slice(0, p.lastIndexOf('/')); }
  function resolve(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = (base ? base.split('/') : []).concat(target.split('/')), out = [];
    parts.forEach(function (s) {
      if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s);
    });
    return out.join('/');
  }
  function relsOf(files, partPath, cache) {
    if (cache[partPath]) return cache[partPath];
    var rp = dirOf(partPath) + '/_rels/' + partPath.slice(partPath.lastIndexOf('/') + 1) + '.rels';
    var m = { byId: {}, list: [] };
    if (hasFile(files, rp)) {
      kids(parse(files[rp]).documentElement, 'Relationship').forEach(function (r) {
        var tg = r.getAttribute('Target'), mode = r.getAttribute('TargetMode');
        var o = { id: r.getAttribute('Id'), type: r.getAttribute('Type') || '', path: mode === 'External' ? null : resolve(dirOf(partPath), tg) };
        m.byId[o.id] = o; m.list.push(o);
      });
    }
    cache[partPath] = m;
    return m;
  }

  // ------------------------------------------------------------ 칸 읽기(python-pptx 동작을 따름)
  var PH_NAMES = { title: 'TITLE', ctrTitle: 'CENTER_TITLE', body: 'BODY', subTitle: 'SUBTITLE', obj: 'OBJECT', tbl: 'TABLE', chart: 'CHART',
    dt: 'DATE', ftr: 'FOOTER', sldNum: 'SLIDE_NUMBER', pic: 'PICTURE', clipArt: 'BITMAP', dgm: 'ORG_CHART', media: 'MEDIA_CLIP',
    sldImg: 'SLIDE_IMAGE', hdr: 'HEADER' };
  var BASE_PH = { BODY: 'BODY', CHART: 'BODY', BITMAP: 'BODY', CENTER_TITLE: 'TITLE', ORG_CHART: 'BODY', DATE: 'DATE', FOOTER: 'FOOTER',
    MEDIA_CLIP: 'BODY', OBJECT: 'BODY', PICTURE: 'BODY', SLIDE_NUMBER: 'SLIDE_NUMBER', SUBTITLE: 'BODY', TABLE: 'BODY', TITLE: 'TITLE' };
  var SHAPE_TAGS = { sp: 1, grpSp: 1, graphicFrame: 1, cxnSp: 1, pic: 1 };

  function shapesOf(tree) { return kids(tree).filter(function (e) { return SHAPE_TAGS[ln(e)]; }); }

  function nvOf(el) {
    var t = ln(el), nv = kid(el, t === 'sp' ? 'nvSpPr' : t === 'pic' ? 'nvPicPr' : t === 'grpSp' ? 'nvGrpSpPr' : t === 'cxnSp' ? 'nvCxnSpPr' : 'nvGraphicFramePr');
    return nv ? kid(nv, 'cNvPr') : null;
  }
  function phEl(sp) {
    var nv = kid(sp, 'nvSpPr'), pr = nv && kid(nv, 'nvPr');
    return pr ? kid(pr, 'ph') : null;
  }
  function phType(sp) {
    var ph = phEl(sp);
    if (!ph) return null;
    var t = attr(ph, 'type') || 'obj';
    return PH_NAMES[t] || 'OBJECT';
  }
  function phIdx(sp) { var v = attr(phEl(sp), 'idx'); return v === null ? 0 : parseInt(v, 10); }

  function xfrmOf(el) {
    var t = ln(el), x;
    if (t === 'graphicFrame') x = kid(el, 'xfrm');
    else if (t === 'grpSp') { var g = kid(el, 'grpSpPr'); x = g && kid(g, 'xfrm'); }
    else { var sp = kid(el, 'spPr'); x = sp && kid(sp, 'xfrm'); }
    return x || null;
  }
  function boxOf(x) {
    if (!x) return null;
    var off = kid(x, 'off'), ext = kid(x, 'ext');
    if (!off || !ext) return null;
    var v = [attr(off, 'x'), attr(off, 'y'), attr(ext, 'cx'), attr(ext, 'cy')].map(function (s) { return s === null ? NaN : Number(s); });
    return v.some(isNaN) ? null : v;
  }

  // 슬라이드 자리표시자가 xfrm 이 없으면 레이아웃 → 마스터 순으로 물려받음
  function inheritedBox(sp, layoutDoc, masterDoc) {
    var own = boxOf(xfrmOf(sp));
    if (own) return own;
    if (!phEl(sp) || !layoutDoc) return null;
    var idx = phIdx(sp), lph = null;
    shapesOf(kid(kid(layoutDoc.documentElement, 'cSld'), 'spTree')).forEach(function (e) {
      if (!lph && ln(e) === 'sp' && phEl(e) && phIdx(e) === idx) lph = e;
    });
    if (!lph) return null;
    var lb = boxOf(xfrmOf(lph));
    if (lb) return lb;
    if (!masterDoc) return null;
    var want = BASE_PH[phType(lph)];
    if (!want) return null;
    var mph = null;
    shapesOf(kid(kid(masterDoc.documentElement, 'cSld'), 'spTree')).forEach(function (e) {
      if (!mph && ln(e) === 'sp' && phEl(e) && phType(e) === want) mph = e;
    });
    return mph ? boxOf(xfrmOf(mph)) : null;
  }

  // 한 칸의 글(문단 사이 \n, 줄바꿈 \v, 필드 글자 포함) — python-pptx text_frame.text
  function frameText(body) {
    if (!body) return '';
    return kids(body, 'p').map(function (p) {
      return kids(p).map(function (e) {
        var t = ln(e);
        if (t === 'r' || t === 'fld') return txt(kid(e, 't'));
        if (t === 'br') return '\x0b';
        return '';
      }).join('');
    }).join('\n');
  }
  // 문단별 글자(a:r 만) — 빈 문단은 뺌
  function paraTexts(body) {
    var out = [];
    if (!body) return out;
    kids(body, 'p').forEach(function (p) {
      var t = kids(p, 'r').map(function (r) { return txt(kid(r, 't')); }).join('').replace(/\x0b/g, '\n');
      if (t.trim()) out.push(t);
    });
    return out;
  }

  // 파이썬 round(x, 3) 와 같게(정확히 반인 값은 짝수 쪽)
  function round3(x) {
    var y = x * 1000;
    if (Math.abs(y % 1) === 0.5) { var f = Math.floor(y); return (f % 2 === 0 ? f : f + 1) / 1000; }
    return Number(x.toFixed(3));
  }

  // 그룹 변환을 풀며 모든 칸을 차례로 (sh, [l,t,w,h], 깊이)
  function walk(tree, cb, T, layoutDoc, masterDoc, depth) {
    T = T || [1, 0, 1, 0];
    shapesOf(tree).forEach(function (sh) {
      var tag = ln(sh), b = (tag === 'sp' ? inheritedBox(sh, layoutDoc, masterDoc) : boxOf(xfrmOf(sh))), l, t, w, h;
      if (b) {
        l = T[1] + b[0] * T[0]; t = T[3] + b[1] * T[2]; w = b[2] * T[0]; h = b[3] * T[2];
      } else { l = t = w = h = 0; }
      cb(sh, [l, t, w, h], depth || 0);
      if (tag === 'grpSp') {
        var x = xfrmOf(sh);
        if (!x) return;
        var ce = kid(x, 'chExt'), co = kid(x, 'chOff');
        var cx = ce ? Number(attr(ce, 'cx')) : 0, cy = ce ? Number(attr(ce, 'cy')) : 0;
        var ox = co ? Number(attr(co, 'x')) : 0, oy = co ? Number(attr(co, 'y')) : 0;
        var nsx = cx ? w / cx : 1, nsy = cy ? h / cy : 1;
        walk(sh, cb, [nsx, l - ox * nsx, nsy, t - oy * nsy], layoutDoc, masterDoc, (depth || 0) + 1);
      }
    });
  }

  var HEAD_ZONE = 0.28;

  function slideInfo(no, doc, ctx) {
    var root = doc.documentElement;
    var info = { no: no, hidden: attr(root, 'show') === '0', layout: ctx.layoutName, items: [], tables: [], pictures: [],
      counts: { picture: 0, table: 0, chart: 0, shape: 0, group: 0, object: 0 } };
    var title = null, message = null, SW = ctx.SW, SH = ctx.SH;
    walk(kid(kid(root, 'cSld'), 'spTree'), function (sh, bb) {
      var tag = ln(sh), l = bb[0], t = bb[1], w = bb[2], h = bb[3];
      var top = t / SH;
      if (tag === 'grpSp') { info.counts.group++; return; }
      if (tag === 'pic') {
        info.counts.picture++;
        var pc = nvOf(sh);
        info.pictures.push({ name: pc ? attr(pc, 'name') || '' : '', descr: pc ? attr(pc, 'descr') || '' : '' });
        return;
      }
      if (tag === 'graphicFrame') {
        var gd = descs(sh, 'graphicData')[0], uri = gd ? attr(gd, 'uri') || '' : '';
        var fn = nvOf(sh), fname = fn ? attr(fn, 'name') || '' : '';
        if (/\/drawingml\/2006\/table$/.test(uri)) {
          info.counts.table++;
          var tbl = descs(gd, 'tbl')[0], rows = [];
          kids(tbl, 'tr').forEach(function (tr) {
            rows.push(kids(tr, 'tc').map(function (tc) { return frameText(kid(tc, 'txBody')).replace(/\x0b/g, '\n').trim(); }));
          });
          info.tables.push({ name: fname, rows: rows });
          var tp = round3(top);
          rows.forEach(function (r, ri) {
            r.forEach(function (tx, ci) {
              if (tx) info.items.push({ kind: 'cell', name: fname + ' r' + ri + 'c' + ci, text: tx, top_pct: tp, ph: null,
                frame: fname, r: ri, c: ci, x: l, y: t, w: w, h: h });
            });
          });
        } else if (/\/drawingml\/2006\/chart$/.test(uri)) info.counts.chart++;
        else info.counts.object++;
        return;
      }
      if (tag === 'cxnSp') return;
      // sp
      var body = kid(sh, 'txBody'), full = frameText(body), hasTxt = !!full.trim();
      var ph = phType(sh), isBox = false, nv = kid(sh, 'nvSpPr'), cs = nv && kid(nv, 'cNvSpPr');
      if (cs && attr(cs, 'txBox') === '1') isBox = true;
      if (!(hasTxt && isBox) && !ph) info.counts.shape++;
      if (hasTxt) {
        var tx = paraTexts(body).join('\n');
        if (!tx.trim()) return;
        var cp = nvOf(sh), tp2 = round3(top);
        info.items.push({ kind: 'text', name: cp ? attr(cp, 'name') || '' : '', text: tx, top_pct: tp2, ph: ph,
          shape_id: cp ? attr(cp, 'id') : null, x: l, y: t, w: w, h: h });
        if (ph && ph.indexOf('TITLE') >= 0 && title === null) title = tx;
        else if (ph && ph.indexOf('BODY') === 0 && message === null && tp2 < HEAD_ZONE) message = tx;
      }
    }, null, ctx.layoutDoc, ctx.masterDoc, 0);
    if (title === null || message === null) {
      var head = info.items.filter(function (i) { return i.kind === 'text' && i.top_pct < HEAD_ZONE; })
        .sort(function (a, b) { return a.top_pct - b.top_pct; });
      if (head.length) {
        if (title === null) title = head[0].text;
        if (message === null) {
          var rest = head.filter(function (i) { return i.text !== title && i.text.length >= 12; });
          if (rest.length) {
            var best = rest[0];
            rest.forEach(function (i) { if (i.text.length > best.text.length) best = i; });
            message = best.text;
          }
        }
      }
    }
    info.title = (title || '').replace(/\n/g, ' ').trim();
    info.message = (message || '').replace(/\n/g, ' ').trim();
    var notes = '';
    if (ctx.notesDoc) {
      var nt = null;
      shapesOf(kid(kid(ctx.notesDoc.documentElement, 'cSld'), 'spTree')).forEach(function (e) {
        if (!nt && ln(e) === 'sp' && phEl(e) && (attr(phEl(e), 'type') || 'obj') === 'body') nt = e;
      });
      if (nt) notes = paraTexts(kid(nt, 'txBody')).join('\n');
    }
    info.notes = notes;
    info.all_text = info.items.map(function (i) { return i.text; }).join('\n');
    return info;
  }

  function coreOf(files) {
    var core = { title: '', subject: '', author: '', last_modified_by: '', keywords: '', comments: '', category: '', company: '' };
    if (hasFile(files, 'docProps/core.xml')) {
      var r = parse(files['docProps/core.xml']).documentElement;
      var map = { title: 'title', subject: 'subject', author: 'creator', last_modified_by: 'lastModifiedBy', keywords: 'keywords', comments: 'description', category: 'category' };
      Object.keys(map).forEach(function (k) { core[k] = txt(descs(r, map[k])[0]); });
    }
    if (hasFile(files, 'docProps/app.xml')) {
      var m = /<Company>([\s\S]*?)<\/Company>/.exec(files['docProps/app.xml']);
      core.company = m ? m[1] : '';
    }
    return core;
  }

  // 마스터·레이아웃 글(저작권 문구·로고 이름)
  function masterTexts(files, pres, cache) {
    var out = [];
    function scan(doc, where) {
      walk(kid(kid(doc.documentElement, 'cSld'), 'spTree'), function (sh) {
        var tag = ln(sh);
        if (tag === 'sp') {
          var f = frameText(kid(sh, 'txBody')).trim();
          if (f) out.push({ where: where, name: attr(nvOf(sh), 'name') || '', text: f });
        } else if (tag === 'pic') {
          out.push({ where: where, name: attr(nvOf(sh), 'name') || '', text: '(그림) ' + (attr(nvOf(sh), 'descr') || '') });
        }
      }, null, null, null, 0);
    }
    var lst = kid(pres.documentElement, 'sldMasterIdLst');
    (lst ? kids(lst, 'sldMasterId') : []).forEach(function (e) {
      var rel = relsOf(files, 'ppt/presentation.xml', cache).byId[attr(e, 'r:id') || attr(e, 'id')];
      if (!rel || !rel.path || !hasFile(files, rel.path)) return;
      var mdoc = parse(files[rel.path]);
      scan(mdoc, '마스터');
      var ll = kid(mdoc.documentElement, 'sldLayoutIdLst');
      (ll ? kids(ll, 'sldLayoutId') : []).forEach(function (le) {
        var lr = relsOf(files, rel.path, cache).byId[attr(le, 'r:id') || attr(le, 'id')];
        if (!lr || !lr.path || !hasFile(files, lr.path)) return;
        var ldoc = parse(files[lr.path]);
        scan(ldoc, '레이아웃 ' + (attr(kid(ldoc.documentElement, 'cSld'), 'name') || ''));
      });
    });
    return out;
  }

  // 전체 추출(파이썬 extract() 와 같은 모양)
  function extractFull(files) {
    var cache = {}, docs = {};
    function doc(p) { return docs[p] || (docs[p] = parse(files[p])); }
    var pres = doc('ppt/presentation.xml');
    var sz = descs(pres.documentElement, 'sldSz')[0];
    var SW = sz ? Number(attr(sz, 'cx')) : 12192000, SH = sz ? Number(attr(sz, 'cy')) : 6858000;
    var prels = relsOf(files, 'ppt/presentation.xml', cache);
    var lst = kid(pres.documentElement, 'sldIdLst'), slides = [];
    (lst ? kids(lst, 'sldId') : []).forEach(function (e, i) {
      var rel = prels.byId[attr(e, 'r:id') || attr(e, 'id')];
      if (!rel || !rel.path || !hasFile(files, rel.path)) return;
      var sr = relsOf(files, rel.path, cache), layoutDoc = null, masterDoc = null, notesDoc = null, layoutName = '';
      sr.list.forEach(function (r) {
        if (!r.path || !hasFile(files, r.path)) return;
        if (/\/slideLayout$/.test(r.type) && !layoutDoc) {
          layoutDoc = doc(r.path);
          layoutName = attr(kid(layoutDoc.documentElement, 'cSld'), 'name') || '';
          relsOf(files, r.path, cache).list.forEach(function (m) {
            if (/\/slideMaster$/.test(m.type) && m.path && hasFile(files, m.path) && !masterDoc) masterDoc = doc(m.path);
          });
        } else if (/\/notesSlide$/.test(r.type) && !notesDoc) notesDoc = doc(r.path);
      });
      slides.push(slideInfo(i + 1, parse(files[rel.path]), { SW: SW, SH: SH, layoutDoc: layoutDoc, masterDoc: masterDoc, notesDoc: notesDoc, layoutName: layoutName }));
    });
    return { slide_count: slides.length, core: coreOf(files), slides: slides, master_texts: masterTexts(files, pres, cache) };
  }

  function extract(files) {
    return extractFull(files).slides.map(function (s) {
      return { no: s.no, layout: s.layout, title: s.title, hidden: s.hidden, notes: s.notes,
        shapes: s.items.map(function (i) { return { name: i.name, text: i.text, kind: i.kind, x: i.x, y: i.y, w: i.w, h: i.h }; }) };
    });
  }

  // ------------------------------------------------------------ 점검 28항목
  var FIX = 'fix', CONFIRM = 'confirm', REF = 'ref';
  var ITEMS = {
    1: ['이전 고객명·다른 회사명 잔재', 1], 2: ['자리표시자·작업 메모', 1], 3: ['회사 소개 수치', 1], 4: ['예시·가설 표시', 1], 5: ['익명 표기 통일', 1],
    6: ['출처·기준일', 0], 7: ['제목 번호·접두', 1], 8: ['목차 간지', 1], 9: ['레이아웃 이름', 1], 10: ['일정 숫자 동기화', 1], 11: ['분량', 0],
    12: ['도식 많음·없음', 0], 13: ['[참고] 분리', 0], 14: ['핵심 문장이 결론형인지', 0], 15: ['단정·방어 유발 문구', 0], 16: ['Why HCG 주장형', 0],
    17: ['종결 어미', 1], 18: ['요청 항목 반영', 0], 19: ['RFP 핵심어', 0], 20: ['고객 표기', 1], 21: ['투입 인력 주요 고객', 0], 22: ['가격 분리', 1],
    23: ['마지막 회의 지시', 0], 24: ['무기명 위반', 1], 25: ['수행 실적 기간', 1], 26: ['지정 목차·필수 장표', 1], 27: ['지정 표지·서식', 1], 28: ['쪽수(본문·요약본)', 0]
  };
  var STD_LAYOUTS = ['표지', '목차', '본문', 'EOD', 'End of document', '끝장', '간지'];
  var LAYOUT_FIX = { '분문': '본문', '본문1__HCG template': '본문', '목차_HCG template': '목차' };
  var NW = '[\\p{L}\\p{N}_]';   // 파이썬 \w (한글 포함)
  var NON_W = /[^\p{L}\p{N}_]/gu;

  function esc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function full(re, s) { return new RegExp('^(?:' + re + ')$').test(s); }
  function snip(text, idx, len, n) { n = n || 18; return text.slice(Math.max(0, idx - n), idx + len + n).replace(/\n/g, ' '); }
  function all(re, s) { return Array.from(s.matchAll(re)); }
  function head(a, n) { return a.slice(0, n); }
  function counter() { return new Map(); }
  function inc(m, k) { m.set(k, (m.get(k) || 0) + 1); }
  function mostCommon(m) {
    var best = null, bn = 0;
    m.forEach(function (v, k) { if (v > bn) { bn = v; best = k; } });
    return best;
  }
  function sum(m) { var s = 0; m.forEach(function (v) { s += v; }); return s; }

  function Ctx(d, o) {
    this.d = d; this.o = o; this.slides = d.slides; this.f = []; this.status = {};
    this.client = o.client || '';
    this.pub = o.publicRules || null;
    this.figures = o.figures || null;
    this.today = o.today || new Date();
    var self = this;
    this.kinds = this.slides.map(function (s, i) { return self.kind(s, i); });
  }
  Ctx.prototype.add = function (item, level, slide, where, msg, suggest, extra) {
    var f = { item: item, name: ITEMS[item][0], level: level, slide: slide || null, shape: where || null, text: msg, how: suggest || '' };
    if (!ITEMS[item][1] && f.how.indexOf('Claude') < 0) f.how = (f.how ? f.how + ' · ' : '') + 'Claude 가 문맥 판단';
    if (extra) Object.keys(extra).forEach(function (k) { f[k] = extra[k]; });
    this.f.push(f);
  };
  Ctx.prototype.kind = function (s, i) {
    var lay = s.layout, t = s.title, n = this.slides.length;
    if (i === 0 || lay.indexOf('표지') >= 0) return 'cover';
    if (i === n - 1 && /EOD|End of|끝장/i.test(lay + t)) return 'eod';
    if (/EOD|End of/i.test(lay)) return 'eod';
    if (lay.indexOf('목차') >= 0 || /^\s*(목차|contents|agenda|table of contents)\s*$/i.test(t)) return 'toc';
    return 'body';
  };
  // 칸별 글 (이름, 글, 위치 정보)
  function textsOf(s, withNotes) {
    var out = s.items.map(function (it) {
      var loc = it.kind === 'cell' ? { cell: { frame: it.frame, r: it.r, c: it.c } } : null;
      return { name: it.name, text: it.text, loc: loc, it: it };
    });
    if (withNotes && s.notes) out.push({ name: '발표자 메모', text: s.notes, loc: null, it: null });
    return out;
  }
  function withLoc(e, extra) {
    var o = {};
    if (e.loc) o.cell = e.loc.cell;
    if (extra) Object.keys(extra).forEach(function (k) { o[k] = extra[k]; });
    return o;
  }

  var FOLD_NAMES = 5;   // 한 장표에 사전 고객명이 이 수 이상이면 장표당 한 건으로 묶음

  function c1(c) {
    if (!c.o.oldNames || !c.o.oldNames.length) { c.status[1] = '입력 없음(oldNames)'; return; }
    var keep = c.client;
    function clean(list) {
      return list.map(function (n) { return String(n).trim(); })
        .filter(function (n) { return n && n.charAt(0) !== '#' && !(keep && (keep.indexOf(n) >= 0 || n.indexOf(keep) >= 0)); });
    }
    var names = clean(c.o.oldNames);
    c.slides.forEach(function (s) {
      var tags = (c.o.slideTags && c.o.slideTags[s.no] && c.o.slideTags[s.no].HCG_OLD) || '';
      var sn = names.slice(), tagNames = clean(String(tags).split('|'));
      tagNames.forEach(function (n) { if (sn.indexOf(n) < 0) sn.push(n); });
      var cases = /\[참고\]|Back-?up|사례|주요\s*고객|Clients?|고객사|레퍼런스|실적/i.test(s.title + s.message);
      var tl = textsOf(s, true).concat(s.pictures.filter(function (p) { return p.descr; }).map(function (p) { return { name: '그림 설명', text: p.descr, loc: null }; }));
      var hits = [];   // 칸 · 사전 이름 · 앞뒤 글 · 읽은 그대로의 글자 · 창고 태그 이름인지
      tl.forEach(function (e) {
        var low = e.text.toLowerCase();
        sn.forEach(function (n) {
          var nl = n.toLowerCase(), idx = -1;
          if (full('[a-z0-9 .&\\-]+', nl)) {
            var m = new RegExp('(?<![a-z0-9])' + esc(nl) + '(?![a-z0-9])').exec(low);
            idx = m ? m.index : -1;
          } else idx = low.indexOf(nl);
          if (idx < 0) return;
          hits.push({ e: e, n: n, sp: snip(e.text, idx, nl.length, 12), orig: e.text.substr(idx, nl.length), tag: tagNames.indexOf(n) >= 0 });
        });
      });
      // 사전 이름이 많은 장표는 주요 고객사·실적·인력 목록일 가능성이 커서 장표당 한 건(확인할 것)으로 묶음.
      // 창고에서 넣은 장의 HCG_OLD 이름은 개수와 상관없이 고칠 것.
      var distinct = [];
      hits.forEach(function (h) { if (!h.tag && distinct.indexOf(h.n) < 0) distinct.push(h.n); });
      var fold = distinct.length >= FOLD_NAMES;
      if (fold) c.add(1, CONFIRM, s.no, '', '고객명 ' + distinct.length + '개(실적 목록일 수 있음): ' + head(distinct, 6).join(', ') + (distinct.length > 6 ? '…' : ''),
        '실적 목록이면 그대로, 이전 제안의 잔재면 바꿈', { names: distinct });
      hits.forEach(function (h) {
        if (fold && !h.tag) return;
        var e = h.e;
        if (cases && !h.tag) c.add(1, CONFIRM, s.no, e.name, '이전 고객명 「' + h.n + '」 — 사례·고객 목록 장표라 의도된 것일 수 있음: …' + h.sp + '…',
          '사례면 A사처럼 익명 표기, 새 고객 장표라면 현 고객명으로', withLoc(e, { from: h.orig }));
        else c.add(1, FIX, s.no, e.name, '이전 고객명 「' + h.n + '」 잔존: …' + h.sp + '…', '현 고객명으로 바꿈',
          withLoc(e, c.client ? { from: h.orig, to: c.client } : { from: h.orig }));
      });
    });
    Object.keys(c.d.core).forEach(function (k) {
      var v = c.d.core[k];
      names.forEach(function (n) {
        if (v && v.toLowerCase().indexOf(n.toLowerCase()) >= 0) c.add(1, FIX, null, '문서 속성 ' + k, '문서 속성에 이전 고객명 「' + n + '」: ' + v, '속성 값 정리');
      });
    });
    c.status[1] = '실행';
  }

  var PLACEHOLDERS = [
    [/(?<![\d,.])0{3}(?![\d,])/g, '「000」'], [/O{3,}|○{3}/g, '「OOO」'],
    [/M{2}\.?\s*D{2}/g, '「MM. DD」'], [/Y{4}\.?\s*M{2}/g, '「YYYY. MM」'], [/(?<![A-Za-z])D{2}(?![A-Za-z])/g, '「DD」'],
    [/(?<![\d.])0\s*주(?!\d)/g, '「0주」'], [/총\s*0\s*개월/g, '「총 0개월」'],
    [/수정\s*중|수정\s*필요|작성\s*중|검토\s*중|TBD|홍길동/g, '작업 메모'],
    [/\[[^\]]*Message[^\]]*\]/g, '「[Message 필요]」'],
    [/소목차[^.\n]{0,20}(?:없는|필요|기입)|기입해\s*주세요|내용을\s*입력|클릭하여|Template Guide/g, '틀 안내문'],
    [/(?:^|\s)(?:MH|HE|CK|YH|ds)\s*(?:수정|확인|작성|검토|추가|중|님)/g, '작성자 이니셜 메모']
  ];

  function c2(c) {
    var anon = !!(c.pub && c.pub['무기명']);
    c.slides.forEach(function (s) {
      textsOf(s, true).forEach(function (e) {
        PLACEHOLDERS.forEach(function (pl) {
          all(pl[0], e.text).forEach(function (m) {
            if ((s.title + e.name).indexOf('인력') >= 0 && full('O{3,}|○{3}', m[0]) && anon) return;
            var lvl = FIX, sug = '삭제하거나 확정값 입력';
            if ((pl[1] === '「OOO」' || pl[1] === '「000」') && /인력|투입|견적|가격|비용|만원/.test(s.title + e.text)) {
              lvl = CONFIRM; sug = '인력 가림·가격 미정 표시일 수 있음 — 의도한 것인지 확인';
            }
            c.add(2, lvl, s.no, e.name, pl[1] + ' 잔존: …' + snip(e.text, m.index, m[0].length) + '…', sug, withLoc(e));
          });
        });
      });
    });
    var cover = c.slides[0];
    if (cover) {
      textsOf(cover, false).forEach(function (e) {
        var m = /(20\d\d)\s*[.\-년]\s*(\d{1,2})/.exec(e.text);
        if (m) {
          var y = parseInt(m[1], 10), mo = parseInt(m[2], 10);
          if (mo >= 1 && mo <= 12 && (c.today.getFullYear() * 12 + c.today.getMonth() + 1) - (y * 12 + mo) > 12)
            c.add(2, CONFIRM, 1, e.name, '표지 날짜 ' + y + '.' + (mo < 10 ? '0' : '') + mo + ' 가 1년 넘게 지남(예전 날짜 잔존?)', '제출일로 갱신', withLoc(e));
        }
      });
    }
    c.status[2] = '실행';
  }

  var FIG_C = [/고객(?:사|기업)?(?:\s*수)?[^\d\n]{0,8}?(\d[\d,]{2,5})\s*(?:여\s*)?(?:개\s*사|개|사|곳|社)/g,
    /(\d[\d,]{2,5})\s*(?:여\s*)?(?:개\s*사|개|사|곳|社)\s*(?:의\s*)?(?:고객|이상의 고객|고객사)/g];
  var FIG_P = [/(?:HR\s*)?(?:전문|컨설팅)?\s*인력[^\d\n]{0,8}?(\d{2,3})\s*명/g, /(\d{2,3})\s*(?:여\s*)?명(?:의)?\s*(?:HR\s*)?(?:전문\s*)?인력/g,
    /(?:HR\s*)?전문가[^\d\n]{0,6}(\d{2,3})\s*명/g];
  var FIG_Y = [/(?:서비스|업력|경험)[^\d\n]{0,6}(\d{2})\s*년/g, /(\d{2})\s*년(?:의|간의)?\s*(?:서비스|업력|경험)/g];
  var FIG_M = /(20\d\d)\s*[.\-년]\s*(\d{1,2})\s*(?:월)?\s*(?:기준|현재)/g;

  // 수치 문구가 든 줄에 붙은 기준월(2026.01 · 2025년 1월 · (2025.01 기준)) → 'YYYY.MM' 또는 null
  function monthNear(t, m) {
    var a = t.lastIndexOf('\n', m.index - 1) + 1, b = t.indexOf('\n', m.index + m[0].length);
    var re = /(20\d\d)\s*(?:[.\-]\s*(\d{1,2})|년\s*(\d{1,2})\s*월?)/g, mm;
    var line = t.slice(a, b < 0 ? t.length : b);
    while ((mm = re.exec(line))) {
      var mo = parseInt(mm[2] || mm[3], 10);
      if (mo >= 1 && mo <= 12) return mm[1] + '.' + (mo < 10 ? '0' : '') + mo;
    }
    return null;
  }

  function figsOf(c) {
    var out = [];
    c.slides.forEach(function (s) {
      var t = s.all_text;
      [['고객수', FIG_C], ['HR인력', FIG_P], ['서비스연수', FIG_Y]].forEach(function (kp) {
        kp[1].forEach(function (p) {
          all(p, t).forEach(function (m) {
            out.push({ no: s.no, k: kp[0], v: parseInt(m[1].replace(/,/g, ''), 10), raw: m[1], mon: monthNear(t, m), snip: snip(t, m.index, m[0].length) });
          });
        });
      });
      all(FIG_M, t).forEach(function (m) {
        var mo = parseInt(m[2], 10);
        out.push({ no: s.no, k: '기준월', v: m[1] + '.' + (mo < 10 ? '0' : '') + mo, raw: m[0], snip: snip(t, m.index, m[0].length) });
      });
    });
    return out;
  }

  function c3(c) {
    var found = figsOf(c), byk = {};
    found.forEach(function (x) { (byk[x.k] = byk[x.k] || new Set()).add(x.v); });
    ['고객수', 'HR인력', '서비스연수'].forEach(function (k) {
      if (byk[k] && byk[k].size > 1) {
        var where = found.filter(function (x) { return x.k === k; }).map(function (x) { return '장표 ' + x.no + ': ' + x.v; }).join('; ');
        c.add(3, FIX, null, k, k + ' 값이 장표마다 다름 (' + where + ')', '한 값으로 통일(수치 표 기준)');
      }
    });
    var fg = c.figures;
    if (fg) {
      var ok = { '고객수': new Set(), 'HR인력': new Set() }, months = new Set();
      (fg['기준'] || []).forEach(function (r) {
        if ((fg['현재_허용'] || []).indexOf(r['기준월']) >= 0) { ok['고객수'].add(r['고객수']); ok['HR인력'].add(r['HR인력']); months.add(r['기준월']); }
      });
      found.forEach(function (x) {
        if (ok[x.k] && ok[x.k].size && !ok[x.k].has(x.v)) {
          var old = fg['기준'].filter(function (r) { return r[x.k] === x.v; }).map(function (r) { return r['기준월']; });
          var okv = Array.from(ok[x.k]).sort(function (a, b) { return a - b; });
          var oldtxt = old.length ? ' · 옛 기준 ' + old.join(',') : '';
          var sug = '수치 표 값(' + okv.join('/') + ')으로 갱신, 기준월 함께 표기 — ※표는 관리자 확인 필요';
          var ft = okv.length === 1 ? { from: x.raw, to: String(okv[0]) } : { from: x.raw };
          // 기준월이 붙은 수치는 공식 기준월과 다르면 옛 기준 값이라 고칠 것, 기준월이 없으면 확인할 것
          if (x.mon) c.add(3, FIX, x.no, x.k, x.k + ' ' + x.v + ' 가 기준월 ' + x.mon + ' 로 표기돼 공식 기준월(' + Array.from(months).sort().join(', ') + ')과 다름(…' + x.snip + '…)' + oldtxt, sug, ft);
          else c.add(3, CONFIRM, x.no, x.k, x.k + ' ' + x.v + ' 가 수치 표와 다름 — 기준월 표기가 없어 옛 제안의 당시 값일 수 있음(…' + x.snip + '…)' + oldtxt, sug, ft);
        }
        if (x.k === '기준월' && months.size && !months.has(x.v))
          c.add(3, CONFIRM, x.no, '기준월', '기준월 ' + x.v + ' 가 수치 표(' + Array.from(months).sort().join(', ') + ')와 다름', '기준월 갱신');
      });
      c.status[3] = '실행(수치 표 대조)' + (fg['관리자_확인_필요'] ? ' · 표는 관리자 확인 필요' : '');
    } else c.status[3] = '일부 실행(장표끼리 비교만, figures 없음)';
  }

  function c4(c) {
    c.slides.forEach(function (s) {
      textsOf(s, false).forEach(function (e) {
        all(/Ex\)|E\.g\.?|Illustrative(?: use only)?|예\)\s/g, e.text).forEach(function (m) {
          c.add(4, FIX, s.no, e.name, '영문·옛 표기 「' + m[0].trim() + '」: …' + snip(e.text, m.index, m[0].length) + '…', '「예시적」·「가설적」으로 통일', withLoc(e));
        });
      });
      var hd = s.title + ' ' + s.message;
      if (/시뮬레이션|Simulation|Output Image|산출물 이미지|Payout|Pay Band/i.test(hd) && !/예시|가설|Illustrative/.test(s.all_text + hd) &&
        (s.counts.table || s.counts.shape > 5))
        c.add(4, CONFIRM, s.no, s.title.slice(0, 20), '가정 수치·산출물 예시로 보이는데 「예시적」·「가설적」 표시가 없음', '제목이나 도식에 「예시적」 표시');
    });
    c.status[4] = '실행';
  }

  function c5(c) {
    var forms = counter(), where = {};
    c.slides.forEach(function (s) {
      textsOf(s, false).forEach(function (e) {
        all(/(?<![A-Za-z가-힣])([A-Z])\s?(사|社)(?![가-힣]{2})/g, e.text).forEach(function (m) {
          inc(forms, m[2]); (where[m[2]] = where[m[2]] || []).push(s.no);
        });
      });
    });
    function sl(k) { return head(Array.from(new Set(where[k] || [])).sort(function (a, b) { return a - b; }), 8); }
    if (forms.size > 1)
      c.add(5, FIX, null, '익명 표기', '「A사」와 「A社」가 섞여 있음 (사 ' + (forms.get('사') || 0) + '곳: 장표 ' + sl('사') + ' / 社 ' + (forms.get('社') || 0) + '곳: 장표 ' + sl('社') + ')', '한 형식(A사)으로');
    c.status[5] = '실행(고객 실명 사례는 1번 항목이 함께 잡음)';
  }

  var NUM_PREFIX = /^\s*((?:[IVX]+|[Ⅰ-Ⅹ])\.|\d+(?:\.\d+)*\.?|\d+\))\s*/;

  function c7(c) {
    var seq = [];
    c.slides.forEach(function (s, i) {
      if (c.kinds[i] !== 'body') return;
      var m = NUM_PREFIX.exec(s.title);
      seq.push([s.no, m ? m[1] : null]);
    });
    if (!seq.length) { c.status[7] = '실행(본문 장표 없음)'; return; }
    var st = { '로마자': 0, '숫자': 0, '없음': 0 };
    seq.forEach(function (x) {
      if (x[1] === null) st['없음']++; else if (/^[IVXⅠ-Ⅹ]/.test(x[1])) st['로마자']++; else st['숫자']++;
    });
    var withP = st['로마자'] + st['숫자'];
    if (withP && st['없음'] >= 0.1 * seq.length) {
      var miss = seq.filter(function (x) { return x[1] === null; }).map(function (x) { return x[0]; });
      c.add(7, FIX, null, '제목 접두', '제목 번호·접두가 있는 장표 ' + withP + '장, 없는 장표 ' + st['없음'] + '장 (없는 장표: ' + head(miss, 12) + ')', '목차 기준으로 접두 일괄 부여');
    }
    if (st['로마자'] && st['숫자'])
      c.add(7, CONFIRM, null, '제목 접두', '로마자·숫자 접두가 섞임 (로마자 ' + st['로마자'] + ' · 숫자 ' + st['숫자'] + ')', '한 형식으로(대분류 로마자·소분류 숫자면 정상)');
    var last = null;
    seq.forEach(function (x) {
      if (x[1] && /^\d+\.$/.test(x[1])) {
        var n = parseInt(x[1].slice(0, -1), 10);
        if (last !== null && n > last + 1) c.add(7, FIX, x[0], '제목 번호', '번호가 ' + last + ' → ' + n + ' 로 건너뜀', '번호 다시 매김');
        last = n;
      }
    });
    c.status[7] = '실행';
  }

  function normW(s) { return s.replace(NON_W, ''); }

  function c8(c) {
    if (c.slides.length < 8) { c.status[8] = '건너뜀(장표 8장 미만)'; return; }
    var toc = [];
    c.kinds.forEach(function (k, i) { if (k === 'toc') toc.push(i); });
    if (!toc.length) {
      c.add(8, CONFIRM, null, '목차', '목차 장표를 못 찾음(레이아웃 이름 「목차」 · 제목 「목차」 기준)', '목차 장표 확인');
      c.status[8] = '실행(목차 없음)';
      return;
    }
    var first = c.slides[toc[0]];
    var entries = first.all_text.split('\n').map(function (x) { return x.trim(); })
      .filter(function (x) { return x.length >= 2 && ['목차', 'contents'].indexOf(x.toLowerCase()) < 0; });
    var nDiv = toc.length - 1;
    var nSec = entries.filter(function (e) { return NUM_PREFIX.test(e) || e.length <= 30; }).length;
    if (nSec >= 3 && nDiv < Math.min(nSec, 3) - 1 && c.slides.length >= 20)
      c.add(8, FIX, null, '목차 간지', '섹션 후보 ' + nSec + '개인데 목차 간지는 ' + toc.length + '장뿐 (본문 ' + c.slides.length + '장)', '섹션이 바뀔 때마다 목차 간지를 넣고 현재 위치 강조');
    var entNorm = entries.map(function (e) { return normW(e.replace(NUM_PREFIX, '')); });
    toc.slice(1).forEach(function (i) {
      var s = c.slides[i];
      var lines = s.all_text.split('\n').filter(function (x) { return x.trim(); }).map(function (x) { return normW(x.trim().replace(NUM_PREFIX, '')); });
      var bad = lines.filter(function (l) { return l && entNorm.length && !entNorm.some(function (e) { return l.indexOf(e) >= 0 || e.indexOf(l) >= 0; }); });
      if (bad.length) c.add(8, CONFIRM, s.no, '간지 항목', '간지 항목명 ' + head(bad, 3).join(', ') + ' 이(가) 첫 목차와 다름', '첫 목차 항목명과 맞춤');
    });
    c.status[8] = '실행';
  }

  function c9(c) {
    var allowed = STD_LAYOUTS.map(function (x) { return x.toLowerCase(); }), seen = new Map();
    c.slides.forEach(function (s) {
      var base = s.layout.replace(/^\d+_/, '').trim().toLowerCase();
      if (allowed.indexOf(base) < 0) { if (!seen.has(s.layout)) seen.set(s.layout, []); seen.get(s.layout).push(s.no); }
    });
    seen.forEach(function (nos, lay) {
      c.add(9, FIX, nos[0], '레이아웃 「' + lay + '」', '표준 밖 레이아웃 이름 「' + lay + '」 (장표 ' + head(nos, 10) + ')',
        '표준 레이아웃 「' + (LAYOUT_FIX[lay] || '표지·목차·본문·끝장 중 하나') + '」로 바꿈', LAYOUT_FIX[lay] ? { from: lay, to: LAYOUT_FIX[lay] } : { from: lay });
    });
    c.status[9] = '실행';
  }

  function weeksOf(c) {
    var out = [];
    c.slides.forEach(function (s) {
      var t = s.all_text;
      [/총\s*(\d{1,2})\s*주/g, /(?<![\d~\-W])(\d{1,2})\s*주\s*(?:간|일정|프로젝트|동안|수행)/g, /(?<![\d~\-W])(\d{1,2})\s*주\s*[·ㆍ,]\s*보고/g].forEach(function (p) {
        all(p, t).forEach(function (m) { out.push([s.no, parseInt(m[1], 10), snip(t, m.index, m[0].length)]); });
      });
      s.tables.forEach(function (tb) {
        var ws = [];
        tb.rows.forEach(function (r) { r.forEach(function (cc) { var m = /^W(\d{1,2})$/.exec(cc.trim()); if (m) ws.push(parseInt(m[1], 10)); }); });
        if (ws.length >= 4) out.push([s.no, Math.max.apply(null, ws), '간트 표 W1~W' + Math.max.apply(null, ws)]);
      });
    });
    return out;
  }

  function c10(c) {
    var w = weeksOf(c), vals = counter();
    w.forEach(function (x) { inc(vals, x[1]); });
    if (vals.size > 1) {
      var parts = [];
      vals.forEach(function (_, v) {
        var nos = Array.from(new Set(w.filter(function (x) { return x[1] === v; }).map(function (x) { return x[0]; }))).sort(function (a, b) { return a - b; });
        parts.push(v + '주 ← 장표 ' + head(nos, 8));
      });
      c.add(10, FIX, null, '일정 주수', '일정 주수가 장표마다 다름: ' + parts.join('; '), '기간 값 하나로 목차·전체 흐름·간트·본문 문장을 모두 갱신(일정 옵션 2안이면 의도된 것일 수 있음)');
    }
    if (c.o.rfpWeeks && w.length) {
      var lo = c.o.rfpWeeks[0], hi = c.o.rfpWeeks[1];
      w.forEach(function (x) {
        if (!(x[1] >= lo && x[1] <= hi)) c.add(10, FIX, x[0], '일정 주수', x[1] + '주가 RFP 기간 ' + lo + '~' + hi + '주 밖(…' + x[2] + '…)', 'RFP 기간 안으로');
      });
    }
    c.status[10] = '실행' + (w.length ? '' : '(주수 표기를 못 찾음)');
  }

  var TAIL_NOUN = /(?:함|임|됨|음|함\)|임\)|있음|없음|것임|예정임|필요함|가능함)[.\s]*$/;
  var TAIL_DA = /(?:다|한다|이다|된다|있다)\s*\.?$/;
  var TAIL_POLITE = /(?:습니다|입니다|합니다|드립니다|세요|해요|예요|이에요)\s*\.?$/;

  var GREET_TAIL = /(?:감사합니다|감사드립니다|드립니다|안녕하십니까|안녕하세요)\s*[.!]?$/;
  var GREET_SLIDE = /인사말|Greeting|Thank\s*you|감사합니다|감사의\s*말/i;

  // 줄 끝이 따옴표 안이면 true(인용문)
  function inQuote(s) {
    var inq = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if ('“「『‘'.indexOf(ch) >= 0) inq = true;
      else if ('”」』’'.indexOf(ch) >= 0) inq = false;
      else if (ch === '"') inq = !inq;
    }
    return inq;
  }

  // 목표 체는 「~함」: 습니다체 문장은 확인할 것(인사말·인용문 제외), 문서 전체가 습니다체면 한 건으로 묶음
  function c17(c) {
    var cnt = counter(), ex = { '습니다체': [], '~다체': [], '하겠음': [] };
    c.slides.forEach(function (s, i) {
      if (c.kinds[i] !== 'body' || GREET_SLIDE.test(s.title + ' ' + s.message)) return;
      var cand = s.message ? [s.message] : [];
      s.items.forEach(function (it) {
        if (it.kind === 'text' && it.ph === null && it.text.length > 25) cand = cand.concat(it.text.split('\n').filter(function (l) { return l.length > 20; }));
      });
      cand.forEach(function (l) {
        l = l.trim();
        if (TAIL_POLITE.test(l)) {
          if (GREET_TAIL.test(l) || inQuote(l)) return;
          inc(cnt, '습니다체'); ex['습니다체'].push([s.no, l]);
        }
        else if (TAIL_NOUN.test(l)) { inc(cnt, '~함체'); if (/하겠음$/.test(l)) ex['하겠음'].push([s.no, l]); }
        else if (TAIL_DA.test(l)) { inc(cnt, '~다체'); ex['~다체'].push([s.no, l]); }
      });
    });
    var main = mostCommon(cnt);
    var total = sum(cnt);
    if (total >= 8 && (cnt.get('습니다체') || 0) >= 0.7 * total) {
      c.add(17, CONFIRM, null, '종결 어미', '전체가 습니다체 — ~함 체로 바꿀지 (습니다체 ' + cnt.get('습니다체') + '문장 / 판정 ' + total + '문장)', '「~함」체로 일괄 변환할지 사람이 정함');
    } else {
      ex['습니다체'].forEach(function (x) { c.add(17, CONFIRM, x[0], '종결 어미', '습니다체 문장(목표 「~함」체): …' + x[1].slice(-24), '「~함」으로 통일(인사말·인용이면 그대로)'); });
    }
    if (ex['하겠음'].length) {
      var nos = Array.from(new Set(ex['하겠음'].map(function (x) { return x[0]; }))).sort(function (a, b) { return a - b; });
      c.add(17, CONFIRM, ex['하겠음'][0][0], '종결 어미', '「~하겠음」 ' + ex['하겠음'].length + '곳 (장표 ' + head(nos, 8) + ')', '「~함」·「~할 것임」 중 하나로 통일');
    }
    c.status[17] = '실행(주 어미: ' + (main || '판정할 문장 없음') + ')';
  }

  function c20(c) {
    var terms = c.o.terms;
    if (!terms) { c.status[20] = '입력 없음(terms)'; return; }
    c.slides.forEach(function (s) {
      textsOf(s, false).forEach(function (e) {
        Object.keys(terms).forEach(function (bad) {
          var good = terms[bad];
          all(new RegExp(esc(bad), 'g'), e.text).forEach(function (m) {
            c.add(20, FIX, s.no, e.name, '「' + bad + '」 → 「' + good + '」: …' + snip(e.text, m.index, m[0].length) + '…', '사전대로 「' + good + '」', withLoc(e, { from: bad, to: good }));
          });
        });
      });
    });
    c.status[20] = '실행';
  }

  var MONEY = /\d[\d,.]*\s*(?:억|천만|백만|만)\s*원|₩\s*\d|\d[\d,]*\s*원(?![\p{L}\p{N}_])|0{3}\s*만\s*원/gu;
  var PRICE_CTX = /견적|용역비|수행\s*비용|프로젝트\s*비용|부가세|VAT|Fee|M\/M|가격|제안\s*금액|산출\s*내역/i;

  function c22(c) {
    c.slides.forEach(function (s) {
      var t = s.title + '\n' + s.all_text, ms = all(MONEY, t);
      if (ms.length && PRICE_CTX.test(t)) c.add(22, FIX, s.no, '금액', '본문에 가격·금액 표기: …' + snip(t, ms[0].index, ms[0][0].length) + '…', '본문에서 빼고 「본문 밖(가격 별도 파일)」 표시');
      else if (ms.length) c.add(22, REF, s.no, '금액', '금액 표기(고객 재무 수치일 수 있음): …' + snip(t, ms[0].index, ms[0][0].length) + '…', '가격이면 본문 밖으로');
    });
    c.status[22] = '실행';
  }

  var LEAK_WORDS = ['HCG', '휴먼컨설팅그룹', 'Human Consulting', 'humanconsulting', 'talenx', '휴넬', 'hunel', 'JaDE', 'DNapp'];
  var LEAK_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[a-z]{2,}|https?:\/\/\S+|www\.\S+|©|저작권|Copyright/i;
  var INFER_RE = /고객\s*(?:사\s*)?\d{3,}\s*(?:개|사|곳)|\d{4}\s*년\s*설립|서비스\s*\d{2}\s*년|전문\s*인력\s*\d{2,3}\s*명|재계약률/;

  function c24(c) {
    var r = c.pub;
    if (!r || !r['무기명']) { c.status[24] = '꺼짐(발주처 규정 없음)'; return; }
    var words = LEAK_WORDS.concat(r['금지어'] || [], r['금지_인명'] || []);
    c.slides.forEach(function (s) {
      textsOf(s, true).forEach(function (e) {
        var low = e.text.toLowerCase();
        words.forEach(function (w) {
          var i = low.indexOf(w.toLowerCase());
          if (i >= 0) c.add(24, FIX, s.no, e.name, '무기명 위반 — 금지어 「' + w + '」: …' + e.text.slice(Math.max(0, i - 10)).slice(0, 40).replace(/\n/g, ' ') + '…',
            '지우거나 「제안사」·「당사」로(인력은 OOO·등급·인원수)', withLoc(e, { from: e.text.substr(i, w.length) }));
        });
        var m = LEAK_RE.exec(e.text);
        if (m) c.add(24, FIX, s.no, e.name, '무기명 위반 — 주소·메일·저작권 문구: ' + m[0].slice(0, 40), '삭제', withLoc(e));
        m = INFER_RE.exec(e.text);
        if (m) c.add(24, CONFIRM, s.no, e.name, '회사를 유추할 수 있는 수치·표현 후보: …' + snip(e.text, m.index, m[0].length) + '…', '사람이 정함(지우기·일반화)', withLoc(e));
      });
      s.pictures.forEach(function (p) {
        if (/logo|로고|워터마크|watermark/i.test(p.name + ' ' + p.descr)) c.add(24, FIX, s.no, '그림 ' + p.name, '로고·워터마크로 보이는 그림 「' + (p.descr || p.name) + '」', '삭제(고객사 로고가 아닌지 확인)');
      });
      if (s.hidden) c.add(24, CONFIRM, s.no, '숨긴 장표', '숨긴 장표가 있음 — 제출본에 남아 있으면 안 됨', '삭제');
    });
    c.d.master_texts.forEach(function (m) {
      words.concat(['©', '저작권', 'Copyright']).forEach(function (w) {
        if (m.text.toLowerCase().indexOf(w.toLowerCase()) >= 0) c.add(24, FIX, null, m.where, '무기명 위반 — ' + m.where + '의 「' + m.name + '」 안에 「' + w + '」', '마스터·레이아웃에서 지움');
      });
    });
    Object.keys(c.d.core).forEach(function (k) {
      var v = c.d.core[k];
      words.forEach(function (w) {
        if (v && v.toLowerCase().indexOf(w.toLowerCase()) >= 0) c.add(24, FIX, null, '문서 속성 ' + k, '문서 속성 ' + k + ' 에 「' + w + '」: ' + v, '속성 지움');
      });
    });
    c.status[24] = '실행';
  }

  function c25(c) {
    var r = c.pub;
    if (!r || !r['최근_N년']) { c.status[25] = '꺼짐(발주처 규정 없음)'; return; }
    var n = parseInt(r['최근_N년'], 10), base = parseInt(r['기준연도'] || c.today.getFullYear(), 10), need = parseInt(r['최소_실적건수'] || 0, 10), ok = 0;
    c.slides.forEach(function (s) {
      if (!/실적|수행\s*사례|레퍼런스/.test(s.title + s.message)) return;
      s.tables.forEach(function (tb) {
        tb.rows.forEach(function (row, ri) {
          var ys = [];
          row.forEach(function (cell) { all(/(?<!\d)(20\d\d)(?!\d)/g, cell).forEach(function (m) { ys.push(parseInt(m[1], 10)); }); });
          if (!ys.length) return;
          var y = Math.max.apply(null, ys);
          if (base - y >= n) c.add(25, FIX, s.no, '실적 표 ' + (ri + 1) + '행', y + '년 실적은 최근 ' + n + '개년(' + (base - n + 1) + '~' + base + ') 밖: ' + row.join(' | ').slice(0, 50), '기간 밖 실적을 빼고 유사 실적으로 채움');
          else ok++;
        });
      });
      if (!new RegExp('최근\\s*' + n + '\\s*개?\\s*년|최근\\s*' + n + '\\s*년').test(s.all_text + s.title))
        c.add(25, CONFIRM, s.no, '실적 표', '「최근 ' + n + '개년 유사 사업」 표기가 없음', '표 제목에 표기');
    });
    if (need && ok < need) c.add(25, FIX, null, '실적 건수', '기간 안 실적 ' + ok + '건 < 기준 ' + need + '건(표에서 연도가 읽힌 행 기준)', '유사 실적 추가');
    c.status[25] = '실행';
  }

  function c26(c) {
    var r = c.pub;
    if (!r || !((r['지정_목차'] || []).length || (r['필수_장표'] || []).length)) { c.status[26] = '꺼짐(발주처 규정 없음)'; return; }
    var titles = c.slides.map(function (s) { return [s.no, normW(s.title + s.message)]; }), last = -1;
    (r['지정_목차'] || []).forEach(function (t) {
      var key = normW(t), pos = -1;
      for (var i = 0; i < titles.length; i++) if (titles[i][1].indexOf(key) >= 0) { pos = i; break; }
      if (pos < 0) c.add(26, FIX, null, '지정 목차', '지정 목차 「' + t + '」 에 해당하는 장표 제목이 없음', '문구를 맞추거나 장표 추가');
      else if (pos < last) c.add(26, FIX, titles[pos][0], '지정 목차', '「' + t + '」 가 지정 순서보다 앞에 있음', '평가 항목 순서대로');
      else last = pos;
    });
    (r['필수_장표'] || []).forEach(function (t) {
      var key = normW(t);
      if (!titles.some(function (x) { return x[1].indexOf(key) >= 0; })) c.add(26, FIX, null, '필수 장표', '필수 장표 「' + t + '」 가 없음', '장표 추가');
    });
    c.status[26] = '실행';
  }

  function c27(c) {
    var r = c.pub;
    if (!r || !((r['지정_레이아웃'] || []).length || (r['인력표_열'] || []).length)) { c.status[27] = '꺼짐(발주처 규정 없음)'; return; }
    if ((r['지정_레이아웃'] || []).length) {
      var allowed = r['지정_레이아웃'].map(function (x) { return x.toLowerCase(); });
      c.slides.forEach(function (s) {
        if (allowed.indexOf(s.layout.toLowerCase()) < 0) c.add(27, FIX, s.no, '레이아웃', '지정 서식 밖 레이아웃 「' + s.layout + '」', '발주처 지정 양식으로');
      });
    }
    if ((r['인력표_열'] || []).length) {
      var need = r['인력표_열'].map(normW);
      c.slides.forEach(function (s) {
        if (/투입\s*인력|수행\s*조직|인력\s*(?:구성|현황)/.test(s.title)) {
          s.tables.forEach(function (tb) {
            var hd = (tb.rows[0] || []).map(normW);
            var miss = need.filter(function (x) { return !hd.some(function (h) { return h.indexOf(x) >= 0; }); });
            if (miss.length) c.add(27, FIX, s.no, '인력 표 열', '발주처 서식 열 ' + miss.join(', ') + ' 이(가) 없음', '지정 서식 열로');
          });
        }
      });
    }
    c.status[27] = '실행';
  }

  // ------- 반자동(후보만)
  function c6(c) {
    c.slides.forEach(function (s, i) {
      if (c.kinds[i] !== 'body') return;
      var t = s.all_text, nums = all(/\d[\d,.]*\s*(?:%|억|조|만\s*명|배|CAGR)/g, t).length;
      if ((s.counts.chart || nums >= 4) && !/출처|Source|자료\s*[:：]|기준\)|\(E\)|※/i.test(t))
        c.add(6, CONFIRM, s.no, s.title.slice(0, 20), '수치·차트가 있는데 출처·기준일 각주가 안 보임(수치 ' + nums + '개, 차트 ' + s.counts.chart + ')', '출처·기준일 각주 추가, 추정치는 (E) 표기');
    });
    c.status[6] = '후보(Claude 가 수치 근거 판단)';
  }

  function isRef(s) { return /\[참고\]|\[Back-?up\]|Back-?up|Appendix|부록|별첨/i.test(s.title + ' ' + s.message); }

  function c11(c) {
    var body = c.slides.filter(function (s, i) { return c.kinds[i] === 'body'; });
    var main = body.filter(function (s) { return !isRef(s); }), refs = body.filter(isRef);
    var flow = main.filter(function (s) { return /전체\s*흐름|Overall|Approach|프로세스 개요|Roadmap/i.test(s.title + s.message); });
    c.add(11, REF, null, '장수', '전체 ' + c.slides.length + '장 · 본문 ' + main.length + '장(표지·목차·간지·끝장·[참고] 뺌) · [참고]·Back-up ' + refs.length + '장');
    if (c.o.pages && main.length > c.o.pages)
      c.add(11, CONFIRM, null, '분량', '본문 ' + main.length + '장 vs 지시 ' + c.o.pages + '장 (' + (main.length - c.o.pages) + '장 초과)',
        '줄일 후보: [참고]로 내릴 장표, 전체 흐름 장표 ' + flow.length + '장(' + head(flow.map(function (s) { return s.no; }), 6) + '), 방법 세부·사례');
    c.status[11] = '실행' + (c.o.pages ? '' : '(분량 지시 없음 → 장수만 표시)');
  }

  function c12(c) {
    var textonly = [], heavy = [];
    c.slides.forEach(function (s, i) {
      if (c.kinds[i] !== 'body') return;
      var cn = s.counts, vis = cn.picture + cn.table + cn.chart + cn.shape + cn.object;
      if (vis <= 2 && s.all_text.length > 150) textonly.push(s.no);
      if (cn.shape + cn.picture >= 40) heavy.push([s.no, cn.shape + cn.picture]);
    });
    if (textonly.length) c.add(12, CONFIRM, textonly[0], '글만 있는 장표', '글만 있고 도식·표가 거의 없는 장표 ' + textonly.length + '장: ' + head(textonly, 15), '도식 추가 후보');
    heavy.forEach(function (x) { c.add(12, CONFIRM, x[0], '도식 과다', '그림·도형 ' + x[1] + '개로 많음', '한 장 한 메시지로 단순화 후보'); });
    c.status[12] = '후보';
  }

  function c13(c) {
    var unmarked = [], marked = [], bodyIdx = [];
    c.slides.forEach(function (s, i) {
      if (c.kinds[i] !== 'body') return;
      if (isRef(s)) marked.push(i);
      else {
        bodyIdx.push(i);
        if (/사례|벤치마킹|Benchmarking|Case(?![\p{L}\p{N}_])|Best Practice|레퍼런스|Reference/iu.test(s.title + s.message)) unmarked.push(s.no);
      }
    });
    if (unmarked.length) c.add(13, CONFIRM, unmarked[0], '[참고] 표시', '사례·벤치마킹으로 보이는데 [참고] 표시가 없는 장표 ' + unmarked.length + '장: ' + head(unmarked, 15), '[참고] 표시를 붙이고 본문 뒤로(핵심 요청 내용이면 본문 유지)');
    if (marked.length && bodyIdx.length && Math.min.apply(null, marked) < Math.max.apply(null, bodyIdx)) {
      var mx = Math.max.apply(null, bodyIdx), early = marked.filter(function (i) { return i < mx; }).map(function (i) { return c.slides[i].no; });
      c.add(13, REF, early[0], '[참고] 위치', '[참고] 장표 ' + early.length + '장이 본문 장표 사이에 있음: ' + head(early, 12), '본문 뒤로 옮길지 판단');
    }
    c.status[13] = '후보';
  }

  var NOUN_END = /(현황|이해|개요|방향|방안|체계|구성|내용|소개|종류|사례|절차|프로세스|일정|목적|범위|필요|결과|구조|모델|제도|항목|기준)\s*$/;
  var VERB_END = /(?:(?:함|임|음|됨|있음|없음|필요함|시점|것임|예정|가능|요구됨|요청됨)\s*\.?|다\s*\.?)$/;

  function c14(c) {
    var cand = 0;
    c.slides.forEach(function (s, i) {
      if (c.kinds[i] !== 'body') return;
      var m = s.message;
      if (!m) {
        if (s.all_text.length > 80) { c.add(14, CONFIRM, s.no, s.title.slice(0, 20), '핵심 문장(제목 아래 한 줄)이 없음', '결론형 문장 후보 2개'); cand++; }
        return;
      }
      if (m.length < 18 || (NOUN_END.test(m) && !VERB_END.test(m))) {
        c.add(14, CONFIRM, s.no, '핵심 문장', '명사 나열·짧은 문장으로 보임: 「' + m.slice(0, 50) + '」', '결론형 한 문장 후보 2개(Claude 가 작성)'); cand++;
      }
    });
    c.status[14] = '후보 ' + cand + '건(Claude 가 문장 판단)';
  }

  var ASSERT = ['우려', '불안정', '구식', '리스크 높음', '반드시', '최적의', '최적 ', '문제는', '해야 한다', '해야 함', '비효율', '실패', '심각', '승리', '적극 활용할 예정', '높은 이해도', '과도한'];

  function c15(c) {
    c.slides.forEach(function (s) {
      textsOf(s, false).forEach(function (e) {
        ASSERT.forEach(function (w) {
          var i = e.text.indexOf(w);
          if (i >= 0) c.add(15, CONFIRM, s.no, e.name, '단정·방어 유발 후보 「' + w.trim() + '」: …' + e.text.slice(Math.max(0, i - 15), i + 20).replace(/\n/g, ' ') + '…', '사실 진술형·「~필요한 시점」으로(Claude 가 문맥 판단)', withLoc(e));
        });
      });
    });
    c.status[15] = '후보';
  }

  var CLAIM = /성공적인[^.\n]{0,20}(?:실시|수행|구축)|국내\s*최고|최고|최대|유일|완벽|보장|확실/g;

  function c16(c) {
    c.slides.forEach(function (s) {
      var why = /Why\s*(?:HCG|Us)|왜\s*HCG|차별/i.test(s.title + s.message);
      textsOf(s, false).forEach(function (e) {
        all(CLAIM, e.text).forEach(function (m) {
          if (why || /HCG|당사/.test(e.text))
            c.add(16, CONFIRM, s.no, e.name, '결과 단정·최상급 후보: …' + snip(e.text, m.index, m[0].length) + '…', '「~경험을 보유함」·사례형으로(국내 최고·최대는 근거 수치와 함께)', withLoc(e));
        });
      });
    });
    c.status[16] = '후보';
  }

  function c18(c) {
    var reqs = c.o.requests;
    if (!reqs) { c.status[18] = '입력 없음(requests)'; return; }
    reqs.forEach(function (r) {
      if (r['본문밖']) return;
      var kws = r.keywords || (r['항목'].match(/[가-힣A-Za-z]{2,}/g) || []), hit = [];
      c.slides.forEach(function (s, i) {
        if (c.kinds[i] !== 'body') return;
        var hay = (s.title + ' ' + s.message).toLowerCase();
        if (kws.some(function (k) { return hay.indexOf(k.toLowerCase()) >= 0; })) hit.push(s);
        else if (kws.some(function (k) { return s.all_text.toLowerCase().indexOf(k.toLowerCase()) >= 0; })) hit.push(s);
      });
      if (!hit.length) c.add(18, CONFIRM, null, '요청 ' + (r.no || ''), '요청 「' + r['항목'] + '」에 반영된 장표가 없음', '장표 묶음 추천(Claude 가 요청 유형→장표 묶음 표에서 고름)');
    });
    c.status[18] = '후보(판정은 Claude 가 확정)';
  }

  function c19(c) {
    var kws = c.o.rfpKeywords;
    if (!kws || !kws.length) { c.status[19] = '입력 없음(rfpKeywords)'; return; }
    var purpose = c.slides.filter(function (s) { return /목적|범위|방향|목표|Why/.test(s.title + s.message); });
    var ptxt = purpose.map(function (s) { return s.title + s.message + s.all_text; }).join(' ');
    var alltxt = c.slides.map(function (s) { return s.all_text + s.title; }).join(' ');
    kws.forEach(function (kw) {
      if (ptxt.indexOf(kw) >= 0) return;
      c.add(19, CONFIRM, null, 'RFP 핵심어', 'RFP 핵심어 「' + kw + '」 가 목적·방향 장표에 없음(' + (alltxt.indexOf(kw) >= 0 ? '다른 장표에는 있음' : '제안서 어디에도 없음') + ')', '목적·방향 장표에 핵심어 배치');
    });
    c.status[19] = '실행(후보)';
  }

  function c21(c) {
    var n = 0;
    c.slides.forEach(function (s) {
      if (!/투입\s*인력|수행\s*조직|인력\s*(?:구성|현황|이력)|Profile/i.test(s.title + s.message)) return;
      s.tables.forEach(function (tb) {
        var hd = tb.rows[0] || [];
        hd.forEach(function (h, i) {
          if (!/주요\s*고객|Clients?|수행\s*실적|주요\s*프로젝트/i.test(h)) return;
          var cells = tb.rows.slice(1).filter(function (r) { return r.length > i && r[i]; }).map(function (r) { return r[i]; });
          if (cells.length) {
            n++;
            c.add(21, CONFIRM, s.no, '인력 표 「' + hd[i] + '」 열', '주요 고객 목록 후보: ' + head(cells, 4).map(function (x) { return x.replace(/\n/g, ', ').slice(0, 40); }).join(' / '), '고객 업종과 맞는 실적을 앞으로(Claude 가 업종 관련성 판단)');
          }
        });
      });
    });
    c.status[21] = '후보 ' + n + '건';
  }

  function c23(c) {
    var items = c.o.meeting;
    if (!items) { c.status[23] = '입력 없음(meeting)'; return; }
    var alltxt = c.slides.map(function (s) { return s.all_text + ' ' + s.title; }).join(' ');
    items.forEach(function (it) {
      var kws = it.keywords || (it['지시'].match(/[가-힣A-Za-z0-9]{2,}/g) || []);
      if (!kws.some(function (k) { return alltxt.indexOf(k) >= 0; })) c.add(23, CONFIRM, null, '회의 지시 ' + (it.id || ''), '지시 「' + it['지시'] + '」 의 키워드가 장표에 없음', '반영 확인표에 미반영 표시');
    });
    c.status[23] = '후보(반영 확인표는 보고서 끝)';
  }

  function c28(c) {
    var r = c.pub;
    if (!r || !(r['본문_쪽수'] || r['요약본_쪽수'] || r['발표_분'])) { c.status[28] = '꺼짐(발주처 규정 없음)'; return; }
    var body = c.slides.filter(function (s, i) { return c.kinds[i] === 'body' && !isRef(s); });
    var lim = r['이_파일은_요약본'] ? r['요약본_쪽수'] : r['본문_쪽수'];
    if (lim && c.slides.length > lim) c.add(28, CONFIRM, null, '쪽수', '전체 ' + c.slides.length + '장이 상한 ' + lim + '쪽을 넘음(본문 ' + body.length + '장)', '넘는 장수만큼 줄일 후보([참고]·흐름 장표) 제시');
    if (r['발표_분']) {
      var per = body.length / Number(r['발표_분']);
      if (per > 2.5) c.add(28, CONFIRM, null, '발표 시간', '본문 ' + body.length + '장 / 발표 ' + r['발표_분'] + '분 = 분당 ' + per.toFixed(1) + '장(빠름)', '발표 분량에 맞춰 압축 후보');
    }
    c.status[28] = '실행(후보)';
  }

  var RUNNERS = [c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, c17, c18, c19, c20, c21, c22, c23, c24, c25, c26, c27, c28];

  function run(files, opts) {
    var c = new Ctx(extractFull(files), opts || {});
    RUNNERS.forEach(function (fn) { fn(c); });
    return { findings: c.f, status: c.status };
  }

  return { init: init, extract: extract, extractFull: extractFull, run: run, ITEMS: ITEMS };
});
