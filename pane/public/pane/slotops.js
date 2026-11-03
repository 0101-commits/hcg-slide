/* HCG 장표 칸 엔진 — 장표 XML 문자열만 다룸(Office.js · JSZip 몰라도 됨).
 * 파이썬 판(hcg-restructure/scripts/pptx_ops.py)의 규칙을 그대로 옮김.
 * 브라우저: window.HCGSlot · Node: module.exports. DOMParser/XMLSerializer 는 init 으로 넣을 수 있음.
 *
 * 공통 옵션 opts = { slideW, slideH }  (EMU. presentation.xml 의 sldSz. 없으면 16:9 12192000x6858000)
 *   ※ 칸 번호 순서가 장표 크기에 따라 달라지므로, 같은 장표엔 늘 같은 크기를 넘길 것.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HCGSlot = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DP = null, XS = null;
  var DEF_W = 12192000, DEF_H = 6858000;

  function init(o) {
    o = o || {};
    if (o.DOMParser) DP = o.DOMParser;
    if (o.XMLSerializer) XS = o.XMLSerializer;
  }

  function parse(xml) {
    var P = DP || (typeof DOMParser !== 'undefined' ? DOMParser : null);
    if (!P) throw new Error('DOMParser 가 없음: HCGSlot.init({DOMParser, XMLSerializer}) 를 먼저 부를 것');
    var doc = new P().parseFromString(xml, 'application/xml');
    var err = doc.getElementsByTagName('parsererror');
    if (err && err.length) throw new Error('XML 해석 실패: ' + err[0].textContent.slice(0, 120));
    if (!doc.documentElement) throw new Error('XML 해석 실패: 빈 문서');
    return doc;
  }

  function serialize(doc, orig) {
    var S = XS || (typeof XMLSerializer !== 'undefined' ? XMLSerializer : null);
    if (!S) throw new Error('XMLSerializer 가 없음');
    var out = new S().serializeToString(doc);
    var m = /^\s*<\?xml[^>]*\?>\s*/.exec(orig);
    if (m && !/^\s*<\?xml/.test(out)) out = m[0].replace(/\s+$/, '') + '\r\n' + out;
    return out;
  }

  // ------------------------------------------------------------ DOM 보조
  function kids(el, name) {
    var r = [], n = el.childNodes;
    for (var i = 0; i < n.length; i++) {
      var c = n[i];
      if (c.nodeType === 1 && (!name || c.localName === name)) r.push(c);
    }
    return r;
  }
  function child(el, name) {
    if (!el) return null;
    var n = el.childNodes;
    for (var i = 0; i < n.length; i++) if (n[i].nodeType === 1 && n[i].localName === name) return n[i];
    return null;
  }
  function desc(el, name, out) {
    out = out || [];
    var n = el.childNodes;
    for (var i = 0; i < n.length; i++) {
      var c = n[i];
      if (c.nodeType !== 1) continue;
      if (c.localName === name) out.push(c);
      desc(c, name, out);
    }
    return out;
  }
  function attr(el, n) { return el && el.hasAttribute(n) ? el.getAttribute(n) : null; }
  function num(el, n, d) {
    var a = attr(el, n);
    if (a === null) return d;
    var v = parseFloat(a);
    return isFinite(v) ? v : d;
  }
  function setNum(el, n, v) { el.setAttribute(n, String(Math.round(v))); }

  var SHAPE_TAGS = { sp: 1, grpSp: 1, graphicFrame: 1, pic: 1, cxnSp: 1 };
  function shapesOf(parent) {
    var r = [], n = parent.childNodes;
    for (var i = 0; i < n.length; i++) if (n[i].nodeType === 1 && SHAPE_TAGS[n[i].localName]) r.push(n[i]);
    return r;
  }

  function spTreeOf(doc) {
    var c = child(doc.documentElement, 'cSld');
    var t = c && child(c, 'spTree');
    if (!t) throw new Error('장표 XML 에 spTree 가 없음');
    return t;
  }

  function nvOf(sh) {
    var nv = null, k = kids(sh);
    for (var i = 0; i < k.length; i++) if (/^nv.*Pr$/.test(k[i].localName)) { nv = k[i]; break; }
    return nv;
  }
  function cNvPrOf(sh) { return child(nvOf(sh), 'cNvPr'); }
  function phType(sh) {
    var nv = nvOf(sh), p = nv && child(child(nv, 'nvPr'), 'ph');
    return p ? (attr(p, 'type') || 'body') : null;
  }

  function xfrmEl(sh) {
    var ln = sh.localName, c;
    if (ln === 'graphicFrame') return child(sh, 'xfrm');
    if (ln === 'grpSp') { c = child(sh, 'grpSpPr'); return c && child(c, 'xfrm'); }
    c = child(sh, 'spPr');
    return c && child(c, 'xfrm');
  }
  function geom(sh, T) {
    var x = xfrmEl(sh);
    if (!x) return null;
    var off = child(x, 'off'), ext = child(x, 'ext');
    if (!off || !ext) return null;
    return {
      l: T.ox + num(off, 'x', 0) * T.sx, t: T.oy + num(off, 'y', 0) * T.sy,
      w: num(ext, 'cx', 0) * T.sx, h: num(ext, 'cy', 0) * T.sy
    };
  }
  function childT(grp, g) {
    var x = xfrmEl(grp), co = child(x, 'chOff'), ce = child(x, 'chExt');
    if (!co || !ce) return null;
    var cx = num(ce, 'cx', 0), cy = num(ce, 'cy', 0);
    var nsx = cx ? g.w / cx : 1, nsy = cy ? g.h / cy : 1;
    return { sx: nsx, ox: g.l - num(co, 'x', 0) * nsx, sy: nsy, oy: g.t - num(co, 'y', 0) * nsy };
  }

  // 모든 도형을 장표 좌표로 훑음(묶음 안 포함)
  function collect(spTree) {
    var out = [];
    (function rec(parent, T, depth, path) {
      shapesOf(parent).forEach(function (sh) {
        var g = geom(sh, T);
        var node = { el: sh, g: g, T: T, depth: depth, path: path, parent: parent };
        out.push(node);
        if (sh.localName === 'grpSp' && g) {
          var nT = childT(sh, g);
          if (nT) rec(sh, nT, depth + 1, path.concat(attr(cNvPrOf(sh), 'id')));
        }
      });
    })(spTree, { sx: 1, ox: 0, sy: 1, oy: 0 }, 0, []);
    return out;
  }

  // ------------------------------------------------------------ 글자 수
  function nsp(s) { return s.replace(/\s/g, '').length; }
  function wlen(s) {
    var w = 0;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i), c = s.charCodeAt(i);
      if (/\s/.test(ch)) w += 0.3;
      else if ((c >= 0xAC00 && c <= 0xD7A3) || (c >= 0x4E00 && c <= 0x9FFF) || ch === '○' || ch === '◯') w += 1;
      else if (/[A-Z]/.test(ch)) w += 0.7;
      else if (/[a-z0-9]/.test(ch)) w += 0.55;
      else w += 0.35;
    }
    return w;
  }
  function r1(v) { return Math.round(v * 10) / 10; }

  function paraText(p) {
    var s = '', k = p.childNodes;
    for (var i = 0; i < k.length; i++) {
      var c = k[i];
      if (c.nodeType !== 1) continue;
      if (c.localName === 'r' || c.localName === 'fld') { var t = child(c, 't'); s += t ? t.textContent : ''; }
      else if (c.localName === 'br') s += '\n';
    }
    return s;
  }
  // txBody → {text, fontPt}
  function bodyInfo(tx) {
    if (!tx) return { text: '', fontPt: null };
    var ps = kids(tx, 'p'), parts = [], fs = null;
    ps.forEach(function (p) {
      parts.push(paraText(p));
      if (fs === null) {
        var rs = kids(p, 'r');
        for (var i = 0; i < rs.length; i++) {
          var rp = child(rs[i], 'rPr'), sz = rp && attr(rp, 'sz');
          if (sz) { fs = Math.round(parseFloat(sz)) / 100; break; }
        }
      }
    });
    return { text: parts.join('\n'), fontPt: fs };
  }

  // ------------------------------------------------------------ 표
  function tblOf(gf) {
    var g = child(gf, 'graphic'), d = g && child(g, 'graphicData'), t = d && child(d, 'tbl');
    return t || null;
  }
  function isSpanned(tc) { return attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1' || attr(tc, 'hMerge') === 'true' || attr(tc, 'vMerge') === 'true'; }

  // ------------------------------------------------------------ 칸 지도
  var DEF_PT = { title: 24, key: 18, head: 14, body: 12, label: 12, note: 9 };

  function guessRole(rec, W, H) {
    var fs = rec.fontPt || 0, wc = rec.wchars, tx = rec.text;
    if (rec.kind === 'cell') return rec.r === 0 ? 'head' : (rec.c === 0 ? 'label' : 'body');
    if (rec.phType === 'title' || rec.phType === 'ctrTitle') return 'title';
    if (rec.y < 0.1 * H && wc <= 22) return 'title';
    if (rec.y < 0.26 * H && wc >= 15 && (fs >= 16 || fs === 0)) return 'key';
    if (rec.y + rec.h > 0.9 * H && ((fs && fs <= 10) || /^\s*(※|\*|출처|1\))/.test(tx))) return 'note';
    if (fs >= 16 && wc <= 25) return 'head';
    if (wc <= 22) return 'head';
    return 'body';
  }

  function shapeKind(sh, hasText) {
    var ln = sh.localName;
    if (ln === 'grpSp') return 'group';
    if (ln === 'graphicFrame') return tblOf(sh) ? 'table' : 'object';
    if (ln === 'pic') return 'pic';
    if (ln === 'cxnSp') return 'line';
    if (hasText) return 'text';
    var sp = child(sh, 'spPr'), pg = sp && child(sp, 'prstGeom'), pr = pg && attr(pg, 'prst');
    if (pr && (pr === 'line' || /Connector/.test(pr))) return 'line';
    return 'shape';
  }

  function insetsOf(bp) {
    return [num(bp, 'lIns', 91440), num(bp, 'tIns', 45720), num(bp, 'rIns', 91440), num(bp, 'bIns', 45720)];
  }

  function slotMap(xml, opts) {
    opts = opts || {};
    var W = opts.slideW || DEF_W, H = opts.slideH || DEF_H;
    var doc = parse(xml), tree = spTreeOf(doc);
    var items = [];
    collect(tree).forEach(function (nd) {
      var sh = nd.el, ph = phType(sh);
      if (ph && /^(sldNum|ftr|dt)$/.test(ph)) return;
      var cn = cNvPrOf(sh), g = nd.g || { l: 0, t: 0, w: 0, h: 0 };
      var tx = sh.localName === 'sp' ? child(sh, 'txBody') : null;
      var info = bodyInfo(tx);
      var hasText = info.text.trim() !== '';
      var kd = shapeKind(sh, hasText);
      if (ph && kd === 'shape') return;
      var rec = {
        no: 0, id: num(cn, 'id', 0), name: attr(cn, 'name') || '', kind: kd, role: null,
        x: Math.round(g.l), y: Math.round(g.t), w: Math.round(g.w), h: Math.round(g.h),
        text: '', chars: 0, wchars: 0, limitLo: 0, limitHi: 0, fontPt: null,
        estLines: 0, overflow: false, tableNo: null, r: null, c: null,
        groupPath: nd.path.join('/') || null
      };
      if (!nd.g) rec.noXfrm = true;
      if (ph) rec.phType = ph;
      if (hasText) fillText(rec, info, tx && child(tx, 'bodyPr'));
      if (kd === 'text') rec.role = guessRole(rec, W, H);
      items.push({ rec: rec, el: sh, table: null });
      if (kd === 'table') {
        var tbl = tblOf(sh);
        var cols = kids(child(tbl, 'tblGrid'), 'gridCol').map(function (c) { return num(c, 'w', 0) * nd.T.sx; });
        var xs = [g.l], ys = [g.t], trs = kids(tbl, 'tr');
        cols.forEach(function (w) { xs.push(xs[xs.length - 1] + w); });
        trs.forEach(function (tr) { ys.push(ys[ys.length - 1] + num(tr, 'h', 0) * nd.T.sy); });
        rec.rows = trs.length; rec.cols = cols.length;
        trs.forEach(function (tr, r) {
          kids(tr, 'tc').forEach(function (tc, c) {
            if (isSpanned(tc)) return;
            var ci = bodyInfo(child(tc, 'txBody'));
            if (!ci.text.trim()) return;
            var cs = Math.max(1, num(tc, 'gridSpan', 1)), rs = Math.max(1, num(tc, 'rowSpan', 1));
            var x2 = xs[Math.min(c + cs, cols.length)], y2 = ys[Math.min(r + rs, trs.length)];
            var cr = {
              no: 0, id: rec.id, name: rec.name, kind: 'cell', role: null,
              x: Math.round(xs[c]), y: Math.round(ys[r]), w: Math.round(x2 - xs[c]), h: Math.round(y2 - ys[r]),
              text: '', chars: 0, wchars: 0, limitLo: 0, limitHi: 0, fontPt: null,
              estLines: 0, overflow: false, tableNo: null, r: r, c: c, rowSpan: rs, colSpan: cs,
              groupPath: rec.groupPath
            };
            var pr = child(tc, 'tcPr');
            fillText(cr, ci, null, [num(pr, 'marL', 91440), num(pr, 'marT', 45720), num(pr, 'marR', 91440), num(pr, 'marB', 45720)]);
            cr.role = guessRole(cr, W, H);
            items.push({ rec: cr, el: tc, table: rec });
          });
        });
      }
    });
    // 위→아래(45px 띠) · 왼→오. 좌표 px = 장표 가로 1600 기준(파이썬 판과 같음)
    var k = 1600 / W;
    items.forEach(function (it, i) { it.i = i; });
    items.sort(function (a, b) {
      var ba = Math.round(Math.round(a.rec.y * k) / 45), bb = Math.round(Math.round(b.rec.y * k) / 45);
      return ba - bb || (Math.round(a.rec.x * k) - Math.round(b.rec.x * k)) || (a.i - b.i);
    });
    items.forEach(function (it, n) { it.rec.no = n + 1; });
    var tabNo = new Map();
    items.forEach(function (it) { if (it.rec.kind === 'table') tabNo.set(it.rec, it.rec.no); });
    items.forEach(function (it) { if (it.table) it.rec.tableNo = tabNo.get(it.table); });
    return items.map(function (it) {
      var s = it.rec;
      if (s.text) { var e = estimateOverflow(s); s.estLines = e.estLines; s.overflow = e.overflow; s.estimated = true; }
      return s;
    });
  }

  function fillText(rec, info, bodyPr, insets) {
    rec.text = info.text;
    rec.chars = nsp(info.text);
    rec.wchars = r1(wlen(info.text));
    rec.limitLo = r1(rec.wchars * 0.8);
    rec.limitHi = r1(rec.wchars * 1.2);
    rec.fontPt = info.fontPt;
    rec.insets = insets || insetsOf(bodyPr);
    if (bodyPr && attr(bodyPr, 'wrap') === 'none') rec.noWrap = true;
  }

  // ------------------------------------------------------------ 넘침 어림
  // 한 줄 폭 ≈ 폭 가중 글자 수 × 글꼴pt × 0.0139in. 줄 수 × 글꼴 × 1.2 > 칸 높이면 넘침. (어림)
  function estimateOverflow(slot) {
    var role = slot.role || 'body';
    var pt = slot.fontPt || DEF_PT[role] || 12;
    var ins = slot.insets || [91440, 45720, 91440, 45720];
    var wIn = (slot.w - ins[0] - ins[2]) / 914400, hIn = slot.h / 914400;
    var cap = wIn > 0 ? wIn / (pt * 0.0139) : 0;   // 한 줄에 들어가는 폭 가중 글자 수
    var lines = 0;
    String(slot.text || '').split('\n').forEach(function (p) {
      if (slot.noWrap || cap <= 0) { lines += 1; return; }
      lines += Math.max(1, Math.ceil(wlen(p) / (cap * 0.95)));
    });
    var needIn = lines * pt * 1.2 / 72 + (ins[1] + ins[3]) / 914400;
    return {
      estLines: lines, fontPtUsed: pt, needH: Math.round(needIn * 914400),
      overflow: !!slot.text && slot.w > 0 && slot.h > 0 && needIn > hIn, estimated: true
    };
  }

  // ------------------------------------------------------------ 겹침
  function overlaps(slots) {
    var ts = slots.filter(function (s) { return (s.kind === 'text' || s.kind === 'cell') && s.w > 0 && s.h > 0; });
    var out = [];
    for (var i = 0; i < ts.length; i++) for (var j = i + 1; j < ts.length; j++) {
      var a = ts[i], b = ts[j];
      if (a.kind === 'cell' && b.kind === 'cell' && a.tableNo === b.tableNo) continue;
      var iw = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
      var ih = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
      if (iw <= 0 || ih <= 0) continue;
      var inter = iw * ih, small = Math.min(a.w * a.h, b.w * b.h);
      if (inter >= 0.97 * small) continue;          // 한 칸이 다른 칸을 감쌈
      if (inter > 0.15 * small) out.push({ a: a.no, b: b.no, area: Math.round(inter), ratio: Math.round(inter / small * 100) / 100 });
    }
    return out;
  }

  // ------------------------------------------------------------ 열 찾기
  function unitHasText(el) {
    var tx = el.localName === 'sp' ? child(el, 'txBody') : null;
    return !!tx && bodyInfo(tx).text.trim() !== '';
  }
  function isSmallConn(nd, W, H) {   // 열 사이 연결 기호 후보: 글 없는 작은 도형·그림
    var b = nd.bb;
    if (b.w > 0.05 * W || b.h > 0.1 * H || b.w <= 0) return false;
    if (nd.el.localName === 'grpSp') return !desc(nd.el, 't').some(function (t) { return t.textContent.trim(); });
    return !unitHasText(nd.el);
  }
  function compOf(c) {
    var s = { text: 0, pic: 0 };
    c.units.forEach(function (u) {
      if (unitHasText(u.el) || u.el.localName === 'graphicFrame') s.text++;
      if (u.el.localName === 'pic') s.pic++;
    });
    return s;
  }
  // 열 폭 ±15% · 구성(글 · 그림 개수) 비슷해야 열 구조. 아니면 이유를 돌려줌.
  function similarity(cols) {
    var ws = cols.map(function (c) { return c.r - c.l; }).sort(function (a, b) { return a - b; });
    var med = ws[Math.floor(ws.length / 2)];
    for (var i = 0; i < cols.length; i++) {
      var w = cols[i].r - cols[i].l;
      if (Math.abs(w - med) > 0.15 * med) return '열 폭이 서로 다름(' + (i + 1) + '번째 열 ' + Math.round(w / med * 100) + '%) — 열 구조가 아님';
    }
    var cs = cols.map(compOf);
    for (var a = 0; a < cs.length; a++) for (var b = a + 1; b < cs.length; b++) {
      var lo = Math.min(cs[a].text, cs[b].text) + 1, hi = Math.max(cs[a].text, cs[b].text) + 1;
      if (lo / hi < 0.3 || Math.abs(cs[a].pic - cs[b].pic) > 3)
        return '열 구성(글 ' + cs[a].text + '·' + cs[b].text + '개, 그림 ' + cs[a].pic + '·' + cs[b].pic + '개)이 서로 다름 — 열 구조가 아님';
    }
    return null;
  }

  // 돌려줌: {ok:false, reason} 또는 {ok:true, cols, seps, conns, span, gap}
  function detectColumns(doc, opts) {
    opts = opts || {};
    var W = opts.slideW || DEF_W, H = opts.slideH || DEF_H;
    var units = [], seps = [], smalls = [];
    (function visit(parent, T, depth) {
      shapesOf(parent).forEach(function (sh) {
        var g = geom(sh, T);
        if (!g) return;
        var l = g.l, t = g.t, w = g.w, h = g.h;
        if (t < 0.18 * H || t > 0.93 * H) return;
        var nd = { el: sh, bb: { l: l, t: t, w: w, h: h }, dep: depth, T: T, parent: parent };
        if (sh.localName === 'grpSp') {
          if (w >= 0.12 * W) {
            var nT = childT(sh, g);
            if (nT) visit(sh, nT, depth + 1);
          } else if (isSmallConn(nd, W, H)) smalls.push(nd); else units.push(nd);
          return;
        }
        if (w < 0.008 * W && h > 0.08 * H) { seps.push(nd); return; }
        if (w > 0.6 * W || (w < 0.01 * W && h < 0.01 * H)) return;
        if (isSmallConn(nd, W, H)) smalls.push(nd); else units.push(nd);
      });
    })(spTreeOf(doc), { sx: 1, ox: 0, sy: 1, oy: 0 }, 0);
    units.sort(function (a, b) { return a.bb.l - b.bb.l; });
    var cl = [];
    units.forEach(function (u) {
      var l = u.bb.l, r = l + u.bb.w, best = null, bov = 0;
      cl.forEach(function (c) {
        var ov = Math.min(r, c.r) - Math.max(l, c.l);
        if (ov > bov) { best = c; bov = ov; }
      });
      if (best && bov >= 0.5 * Math.min(u.bb.w, best.r - best.l)) {
        best.units.push(u); best.l = Math.min(best.l, l); best.r = Math.max(best.r, r);
      } else cl.push({ l: l, r: r, units: [u] });
    });
    cl = cl.filter(function (c) { return c.units.length >= 3; });
    if (cl.length < 2) return { ok: false, reason: '같은 가로 범위 요소 묶음이 2열 미만 — 열 구조를 못 찾음' };
    cl.sort(function (a, b) { return a.l - b.l; });
    var why = similarity(cl);
    if (why) return { ok: false, reason: why };
    // 작은 도형: 열 가운데 부분이면 그 열의 요소(아이콘 등), 열 사이 경계 부근이면 연결 기호
    var conns = [];
    for (var j = 0; j < cl.length - 1; j++) conns.push([]);
    smalls.forEach(function (s) {
      var cx = s.bb.l + s.bb.w / 2, home = null;
      cl.forEach(function (c) { if (cx >= c.l + 0.1 * (c.r - c.l) && cx <= c.r - 0.1 * (c.r - c.l)) home = c; });
      if (home) { home.units.push(s); return; }
      for (var i = 0; i < cl.length - 1; i++) {
        var mid = (cl[i].r + cl[i + 1].l) / 2, gp = Math.max(0, cl[i + 1].l - cl[i].r);
        if (Math.abs(cx - mid) <= gp / 2 + 0.1 * (cl[i].r - cl[i].l)) { conns[i].push(s); return; }
      }
    });
    var gsum = 0;
    for (var k = 0; k < cl.length - 1; k++) gsum += Math.max(0, cl[k + 1].l - cl[k].r);
    seps.sort(function (a, b) { return a.bb.l - b.bb.l; });
    return {
      ok: true, cols: cl, seps: seps, conns: conns, span: [cl[0].l, cl[cl.length - 1].r],
      gap: gsum / (cl.length - 1)
    };
  }

  function columnsInfo(xml, opts) {
    var det = detectColumns(parse(xml), opts);
    if (!det.ok) return { columns: [], reason: det.reason };
    return {
      reason: null,
      columns: det.cols.map(function (c) {
        return {
          x0: Math.round(c.l), x1: Math.round(c.r),
          members: c.units.map(function (u) { return num(cNvPrOf(u.el), 'id', 0); })
        };
      })
    };
  }
  function columns(xml, opts) { return columnsInfo(xml, opts).columns; }

  // ------------------------------------------------------------ 편집 공통
  function allIds(doc) {
    return desc(doc.documentElement, 'cNvPr').map(function (c) { return num(c, 'id', 0); });
  }
  function maxId(doc) { return allIds(doc).reduce(function (a, b) { return Math.max(a, b); }, 0); }

  function cloneShape(doc, el, ctr, idMap) {
    var cp = el.cloneNode(true);
    desc(cp, 'cNvPr').forEach(function (c) {
      var old = c.getAttribute('id');
      ctr.v += 1;
      c.setAttribute('id', String(ctr.v));
      idMap[old] = String(ctr.v);
    });
    var top = cNvPrOf(cp);
    if (top) top.setAttribute('name', (top.getAttribute('name') || '') + ' (추가)');
    return cp;
  }
  function remapConnectors(cp, idMap) {
    ['stCxn', 'endCxn'].forEach(function (n) {
      desc(cp, n).forEach(function (e) {
        var o = attr(e, 'id');
        if (o !== null && idMap[o]) e.setAttribute('id', idMap[o]);
      });
    });
  }
  function ids(el) {
    var r = {};
    desc(el, 'cNvPr').forEach(function (c) { r[c.getAttribute('id')] = 1; });
    return r;
  }
  // 지운 도형 id 를 가리키는 연결 · 애니메이션 정리
  function cleanDangling(doc, deleted, warnings) {
    if (!Object.keys(deleted).length) return;
    ['stCxn', 'endCxn'].forEach(function (n) {
      desc(doc.documentElement, n).forEach(function (e) {
        if (deleted[attr(e, 'id')] && e.parentNode) e.parentNode.removeChild(e);
      });
    });
    var tm = child(doc.documentElement, 'timing');
    if (tm) {
      var hit = desc(tm, 'spTgt').some(function (e) { return deleted[attr(e, 'spid')]; });
      if (hit) {
        doc.documentElement.removeChild(tm);
        warnings.push('지운 도형을 가리키는 애니메이션이 있어 장표의 애니메이션을 모두 지움');
      }
    }
  }
  function dropEmptyGroups(tree) {
    var changed = true;
    while (changed) {
      changed = false;
      desc(tree, 'grpSp').forEach(function (g) {
        if (g.parentNode && !shapesOf(g).length) { g.parentNode.removeChild(g); changed = true; }
      });
    }
  }
  function depthOf(el) { var d = 0; while (el.parentNode) { d++; el = el.parentNode; } return d; }

  // 묶음의 가로 틀을 자식 bbox 에 다시 맞춤(자식 좌표는 그대로, 묶음 변환만 갱신)
  function fitGroupX(g) {
    var x = xfrmEl(g), off = child(x, 'off'), ext = child(x, 'ext'), co = child(x, 'chOff'), ce = child(x, 'chExt');
    if (!x || !co || !ce) return;
    var cl = Infinity, cr = -Infinity;
    shapesOf(g).forEach(function (k) {
      var kx = xfrmEl(k), ko = kx && child(kx, 'off'), ke = kx && child(kx, 'ext');
      if (!ko || !ke) return;
      cl = Math.min(cl, num(ko, 'x', 0)); cr = Math.max(cr, num(ko, 'x', 0) + num(ke, 'cx', 0));
    });
    if (!isFinite(cl) || cr <= cl) return;
    var cx = num(ce, 'cx', 0), sx = cx ? num(ext, 'cx', 0) / cx : 1;
    setNum(off, 'x', num(off, 'x', 0) + (cl - num(co, 'x', 0)) * sx);
    setNum(ext, 'cx', (cr - cl) * sx);
    setNum(co, 'x', cl);
    setNum(ce, 'cx', cr - cl);
  }

  function setAbs(nd, l, w) {
    var x = xfrmEl(nd.el), off = child(x, 'off'), ext = child(x, 'ext');
    if (l !== null && l !== undefined) setNum(off, 'x', (l - nd.T.ox) / nd.T.sx);
    if (w !== null && w !== undefined) {
      var oldW = num(ext, 'cx', 0), newW = Math.round(w / nd.T.sx);
      setNum(ext, 'cx', newW);
      var tbl = nd.el.localName === 'graphicFrame' ? tblOf(nd.el) : null;   // 표는 열 너비도 같이 비례 조정
      if (tbl && oldW > 0) {
        var gcs = kids(child(tbl, 'tblGrid'), 'gridCol'), acc = 0;
        gcs.forEach(function (gc, i) {
          var nw = i === gcs.length - 1 ? newW - acc : Math.round(num(gc, 'w', 0) * newW / oldW);
          setNum(gc, 'w', nw); acc += nw;
        });
      }
    }
  }

  // 아이콘 = 글 없는 그림·도형 중 열 폭의 25% 이하 · 가로세로비 0.75~1.33
  function isIcon(nd, colW) {
    var b = nd.bb, ln = nd.el.localName;
    if ((ln !== 'pic' && ln !== 'sp') || b.h <= 0 || b.w > 0.25 * colW) return false;
    if (ln === 'sp' && unitHasText(nd.el)) return false;
    var a = b.w / b.h;
    return a >= 0.75 && a <= 1.33;
  }
  // 그림 크기·위치(세로는 가운데 유지)
  function setPic(nd, cxAbs, w, h) {
    var x = xfrmEl(nd.el), off = child(x, 'off'), ext = child(x, 'ext');
    var cyAbs = nd.bb.t + nd.bb.h / 2;
    setNum(off, 'x', (cxAbs - w / 2 - nd.T.ox) / nd.T.sx);
    setNum(off, 'y', (cyAbs - h / 2 - nd.T.oy) / nd.T.sy);
    setNum(ext, 'cx', w / nd.T.sx);
    setNum(ext, 'cy', h / nd.T.sy);
  }

  // 열 사이 간격은 그대로 두고 남은 폭을 N등분
  // order: [{l, w, units}] 새 열 순서 · connB: 경계별 연결 기호 목록(길이 N-1)
  function relayout(doc, order, det, W, parents, st, connB) {
    var L = Math.max(det.span[0], 0), R = Math.min(det.span[1], W), N = order.length;   // 장표 밖으로 삐져나간 배경 도형은 범위에서 뺌
    var g = det.gap, Wn = (R - L - (N - 1) * g) / N;
    if (Wn < 0.3 * (R - L) / N) { g = 0; Wn = (R - L) / N; }
    var left = function (i) { return L + i * (Wn + g); };
    var touched = parents.slice();
    order.forEach(function (col, i) {
      var nl = left(i), ratio = col.w ? Wn / col.w : 1;
      col.units.forEach(function (u) {
        var b = u.bb, off = b.l - col.l;
        if (isIcon(u, col.w)) {
          if (b.w > Wn) { setAbs(u, nl, Wn); }
          else setAbs(u, nl + (Wn - b.w) / 2);
        } else if (u.el.localName === 'pic') {
          var k = Math.min(ratio, 1);   // 그림은 비율 지킴 · 늘릴 땐 원래 크기까지만
          setPic(u, nl + (off + b.w / 2) * ratio, b.w * k, b.h * k);
        } else setAbs(u, nl + off * ratio, b.w * ratio);
        touched.push(u.el.parentNode);
      });
    });
    connB.forEach(function (list, j) {
      var mid = left(j) + Wn + g / 2;
      list.forEach(function (c) { setAbs(c, mid - c.bb.w / 2); touched.push(c.el.parentNode); });
    });
    var seps = det.seps.slice();
    while (seps.length > N - 1) {
      var s = seps.pop();
      touched.push(s.el.parentNode);
      collectIds(s.el, st.deleted);
      s.el.parentNode.removeChild(s.el);
      st.removed += 1;
    }
    while (seps.length < N - 1 && seps.length) {
      var last = seps[seps.length - 1], idMap = {};
      var cp = cloneShape(doc, last.el, st.ctr, idMap);
      remapConnectors(cp, idMap);
      last.el.parentNode.insertBefore(cp, last.el.nextSibling);
      seps.push({ el: cp, bb: last.bb, T: last.T, dep: last.dep });
      st.added += 1;
    }
    seps.forEach(function (s, i) { setAbs(s, left(i) + Wn + g / 2 - s.bb.w / 2); touched.push(s.el.parentNode); });
    var gs = [];
    touched.forEach(function (p) {
      while (p && p.localName === 'grpSp') { if (gs.indexOf(p) < 0) gs.push(p); p = p.parentNode; }
    });
    gs.sort(function (a, b) { return depthOf(b) - depthOf(a); }).forEach(fitGroupX);
    return { Wn: Wn, gap: g };
  }
  function collectIds(el, into) { desc(el, 'cNvPr').forEach(function (c) { into[c.getAttribute('id')] = 1; }); }
  function countIds(el) { return desc(el, 'cNvPr').length; }

  // 겹침 쌍을 (종류·이름·글 앞부분) 열쇠로 바꿔 전후 비교
  function overlapKeys(xml, opts) {
    var map = slotMap(xml, opts), by = {};
    map.forEach(function (s) { by[s.no] = s.kind + '|' + s.name.replace(/ \(추가\)$/, '') + '|' + s.text.slice(0, 30); });
    return overlaps(map).map(function (p) { return [by[p.a], by[p.b]].sort().join('##'); });
  }

  // ------------------------------------------------------------ 열 수 바꾸기
  function setColumns(xml, n, opts) {
    opts = opts || {};
    var W = opts.slideW || DEF_W;
    var log = [], warnings = [];
    var res = { xml: xml, log: log, warnings: warnings, stats: { removed: 0, added: 0 } };
    n = Math.round(n);
    if (!(n >= 2 && n <= 5)) { warnings.push('열 수는 2~5 만 됨(받은 값 ' + n + ')'); return res; }
    var doc = parse(xml), det = detectColumns(doc, opts);
    if (!det.ok) { warnings.push(det.reason + ' — 바꾸지 않음'); return res; }
    var cur = det.cols.length;
    if (n === cur) { log.push('이미 ' + cur + '열'); return res; }
    var st = { removed: 0, added: 0, deleted: {}, ctr: { v: maxId(doc) } };
    var order = [], parents = [];
    var cols = det.cols.map(function (c) { return { l: c.l, w: c.r - c.l, units: c.units }; });
    var connB = det.conns.map(function (l) { return l.slice(); });
    var dropNode = function (u) {
      parents.push(u.el.parentNode);
      st.removed += countIds(u.el);
      collectIds(u.el, st.deleted);
      u.el.parentNode.removeChild(u.el);
    };
    if (n < cur) {
      for (var i = n; i < cur; i++) {
        cols[i].units.forEach(dropNode);
        log.push((i + 1) + '번째 열 삭제(요소 ' + cols[i].units.length + '개)');
      }
      for (var j = n - 1; j < connB.length; j++) {   // 마지막 경계의 연결 기호도 지움
        connB[j].forEach(dropNode);
        if (connB[j].length) log.push('열 사이 기호 ' + connB[j].length + '개 삭제');
      }
      connB = connB.slice(0, n - 1);
      order = cols.slice(0, n);
      dropEmptyGroups(spTreeOf(doc));
    } else {
      order = cols.slice();
      var lastC = cols[cur - 1], lastConn = connB[cur - 2] || [];
      var cloneNodes = function (list) {
        var ctr = st.ctr, maps = {};
        var out = list.map(function (u) {
          var idMap = {}, cp = cloneShape(doc, u.el, ctr, idMap);
          st.added += countIds(cp);
          u.el.parentNode.insertBefore(cp, u.el.nextSibling);
          for (var o in idMap) maps[o] = idMap[o];
          return { el: cp, bb: u.bb, T: u.T, dep: u.dep };
        });
        out.forEach(function (u) { remapConnectors(u.el, maps); });
        return out;
      };
      for (var k = cur; k < n; k++) {
        order.push({ l: lastC.l, w: lastC.w, units: cloneNodes(lastC.units) });
        connB.push(cloneNodes(lastConn));
        log.push('마지막 열 복제 → ' + (k + 1) + '번째 열(이름 끝 「(추가)」)' + (lastConn.length ? ' · 열 사이 기호 ' + lastConn.length + '개 복제' : ''));
      }
    }
    var lay = relayout(doc, order, det, W, parents, st, connB);
    cleanDangling(doc, st.deleted, warnings);
    log.push(cur + '열 → ' + n + '열, 열 폭 ' + Math.round(lay.Wn) + ' EMU · 열 사이 간격 ' + Math.round(lay.gap) + ' EMU 유지');
    if (n > cur) log.push('복제한 열의 글은 원래 열과 같음 — 새 내용으로 바꿔 쓸 것');
    res.xml = serialize(doc, xml);
    res.stats = { removed: st.removed, added: st.added };
    // 바꾼 뒤 검사: 원본에 없던 글 칸 겹침
    try {
      var had = {};
      overlapKeys(xml, opts).forEach(function (k2) { had[k2] = (had[k2] || 0) + 1; });
      var fresh = overlapKeys(res.xml, opts).filter(function (k2) { if (had[k2]) { had[k2]--; return false; } return true; });
      if (fresh.length) warnings.push('새로 생긴 글 칸 겹침 ' + fresh.length + '쌍 — 눈으로 확인할 것');
    } catch (e) { warnings.push('겹침 검사 실패: ' + e.message); }
    return res;
  }

  // ------------------------------------------------------------ 표 행 삭제
  function deleteTableRow(xml, tableNo, rowIndex, opts) {
    opts = opts || {};
    var warnings = [], res = { xml: xml, warnings: warnings };
    var map = slotMap(xml, opts);
    var ts = null;
    for (var i = 0; i < map.length; i++) if (map[i].kind === 'table' && map[i].no === tableNo) ts = map[i];
    if (!ts) { warnings.push('표 칸 번호 ' + tableNo + ' 을 못 찾음'); return res; }
    var doc = parse(xml), frame = null;
    collect(spTreeOf(doc)).forEach(function (nd) {
      if (nd.el.localName === 'graphicFrame' && tblOf(nd.el) && num(cNvPrOf(nd.el), 'id', -1) === ts.id) frame = frame || nd;
    });
    if (!frame) { warnings.push('표 요소를 못 찾음'); return res; }
    var tbl = tblOf(frame.el), trs = kids(tbl, 'tr'), n = trs.length;
    var rows = (Array.isArray(rowIndex) ? rowIndex : [rowIndex]).slice().sort(function (a, b) { return a - b; });
    var head = opts.headerRows === undefined ? 1 : opts.headerRows;
    if (rows.some(function (r) { return !(r >= 0 && r < n); })) { warnings.push('행 번호 범위 밖(표는 ' + n + '행, 0부터)'); return res; }
    if (rows.some(function (r) { return r < head; })) { warnings.push('머리 행(' + head + '행)은 못 지움'); return res; }
    if (n - rows.length <= head) { warnings.push('본문 행을 모두 지울 수 없음'); return res; }
    var dele = {};
    rows.forEach(function (r) { dele[r] = 1; });
    var all = function (a, b) { for (var q = a; q < b; q++) if (!dele[q]) return false; return true; };
    for (var r = 0; r < n; r++) {
      var tcs = kids(trs[r], 'tc');
      for (var c = 0; c < tcs.length; c++) {
        var tc = tcs[c], span = num(tc, 'rowSpan', 1);
        if (span > 1 && dele[r] && !all(r, r + span)) {
          warnings.push('행 ' + r + ' 은 세로 병합 셀의 일부라 못 지움(병합 범위 전체를 함께 지워야 함)'); return res;
        }
        var vm = attr(tc, 'vMerge');
        if ((vm === '1' || vm === 'true') && dele[r]) {
          var o = r - 1;
          while (o >= 0) {
            var oc = kids(trs[o], 'tc')[c], ov = oc && attr(oc, 'vMerge');
            if (ov === '1' || ov === 'true') o--; else break;
          }
          var so = o >= 0 ? num(kids(trs[o], 'tc')[c], 'rowSpan', 1) : 1;
          if (o < 0 || !all(o, o + so)) { warnings.push('행 ' + r + ' 은 세로 병합 범위의 일부라 못 지움'); return res; }
        }
      }
    }
    var oldH = trs.map(function (t) { return num(t, 'h', 0); });
    var sum = oldH.reduce(function (a, b) { return a + b; }, 0);
    var deletedIds = {};
    rows.forEach(function (rr) { collectIds(trs[rr], deletedIds); tbl.removeChild(trs[rr]); });
    var left = kids(tbl, 'tr'), headH = 0, m = left.length - head;
    for (var h = 0; h < head; h++) headH += oldH[h];
    var body = sum - headH, each = Math.floor(body / m), rem = body - each * m;
    left.forEach(function (t, idx) { if (idx >= head) setNum(t, 'h', each + (idx === left.length - 1 ? rem : 0)); });
    var ext = child(child(frame.el, 'xfrm'), 'ext');
    if (ext) setNum(ext, 'cy', sum);
    cleanDangling(doc, deletedIds, warnings);
    res.xml = serialize(doc, xml);
    res.rows = left.length;
    res.rowH = each;
    return res;
  }

  function slideSize(presentationXml) {
    var m = /<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(presentationXml);
    return m ? { slideW: +m[1], slideH: +m[2] } : { slideW: DEF_W, slideH: DEF_H };
  }

  return {
    init: init, slotMap: slotMap, estimateOverflow: estimateOverflow, overlaps: overlaps,
    columns: columns, columnsInfo: columnsInfo, setColumns: setColumns, deleteTableRow: deleteTableRow, slideSize: slideSize,
    wlen: wlen
  };
});
