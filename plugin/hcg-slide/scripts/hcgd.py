# /hcg-draft 단계 스크립트 — Claude Code 스킬이 단계마다 부름. 결과는 짧은 JSON 한 줄로 찍음.
# 사용: hcgd <단계> <데이터 폴더> <작업 폴더> [더]
#   setup   열쇠 준비(다운로드 폴더의 .hcgkey → 이 PC 잠금) · 창고 목록 받기
#   start   작업 폴더의 input.txt · client.json 을 보고 기준 제안서 후보 3건 · DART 상장 여부 · 업종 목록
#   docs    고객 자료 폴더(RFP · 회의록 · 리포트 등)의 글 → docs.md (원본 폴더는 읽기만)
#   plan    후보 장 좁히기 + 공시 숫자 → plan_prompt.md (계획은 hcg-planner 가 plan.json 으로)
#   build   장 순서 → 장표 받기 · 합치기 → 고쳐 쓰기 요청문 rw_1~5.md (+ 원본 자리 재기를 뒤에서)
#   apply   ans_*.json 넣기 → 넘침 · 겹침 재기 → 줄일 게 있으면 fix_1.md
#   fix     fix_<k>.json 넣기 → 다시 재기 → 남으면(2번까지) fix_<k+1>.md
#   names   다른 고객 이름이 남았으면 names.md
#   finish  names.json 넣기 → 그림 속 글 딱지 → 파일 속성 → 확인 목록 → 결과 파일
import collections, json, os, re, shutil, subprocess, sys, time

S = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, S)
import store  # noqa: E402

PS = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File"]
FRAME = {"표지": 1, "Ⅰ": 2, "Ⅱ": 3, "Ⅲ": 4, "Ⅳ": 5, "끝장": 6}  # frame_std.pptx 장 번호
# HCG 소개(Ⅳ장) — 고쳐 쓰지 않고 그대로 넣는 표준 장(2026.10 판). 4장이면 앞 4장
INTRO = [("s22311", "HCG Overview"), ("s22312", "사업 영역"), ("s22313", "컨설팅 서비스 영역"), ("s22314", "주요 고객사"), ("s22315", "Long-term Relationship")]
# Ⅲ장 핵심 메시지 — 제안서마다 쓰는 통상 문구
FIXED_MSG = {"T302": "본 프로젝트의 성공적 수행을 위하여 최적의 경험과 역량을 갖춘 컨설턴트를 투입하여 팀을 구성할 것임",
             "T303": "본 프로젝트 참여 인력의 주요 이력사항은 아래와 같음",
             "T304": "HCG는 프로젝트 품질제고를 위한 일정/품질/위험/보안관리 및 지식전수/사후관리/실행지원 등을 통해 프로젝트의 성공 및 컨설팅 결과의 실행력 제고에 기여하고자 함"}
III_HEAD = {"T301": "추진 일정", "T302": "수행 조직", "T303": "투입 인력", "T304": "프로젝트 관리방안"}
# 계획 뒤에 저절로 붙이는 HCG 표준 장. ids 는 앞에서부터 — 이번 고객 제안서 것이면 다음 id.
# sys=True 는 요청 범위에 시스템(talenx · SaaS 등)이 있을 때만.
STD_ROWS = [
    {"chap": "Ⅰ", "sys": True, "type": "T408", "ids": ["s22279", "s21682", "s06099", "s02071"],
     "head": "Why HCG    2. 자체 성과관리 솔루션 보유 (1/2)",
     "msg": "HCG는 풍부한 컨설팅 경험과 IT 개발 경험을 바탕으로 성과관리 · 조직 · 인사 등 HR 주요 기능을 통합 제공하는 자체 SaaS 플랫폼 talenx를 개발 · 운영하고 있음"},
    {"chap": "Ⅰ", "sys": True, "type": "T408", "ids": ["s22280"],
     "head": "Why HCG    2. 자체 성과관리 솔루션 보유 (2/2)",
     "msg": "HCG는 컨설팅에서 확정한 평가제도와 운영기준을 talenx에 그대로 구현하여 제도 설계와 시스템 운영 간의 단절을 최소화 함"},
    {"chap": "Ⅰ", "sys": True, "type": "T408", "ids": ["s22281"],
     "head": "[참고] talenx 사용자 편의성 향상",
     "msg": "talenx는 HR 담당자의 운영 편의성과 사용자 경험을 함께 고려하여, 평가운영과 참여 全 과정의 편의성을 높인 솔루션임"},
    {"chap": "Ⅰ", "sys": False, "type": "T406", "ids": ["s22282", "s21685", "s09569", "s21618"],
     "head": "Why HCG    3. 지속적인 변화관리 지원",
     "msg": "프로젝트 결과를 바탕으로 성공적인 제도 개선을 추진할 수 있도록 전담 Hot Line을 구성하여 지속적으로 실행지원-모니터링-사후관리를 실시함"},
    {"chap": "Ⅲ", "sys": False, "type": "T302", "ids": ["s22308", "s22265", "s21919", "s21129", "s06139"],
     "head": III_HEAD["T302"], "msg": FIXED_MSG["T302"]},
    {"chap": "Ⅲ", "sys": False, "type": "T303", "ids": ["s22309", "s23555", "s06140", "s09589", "s21658"],
     "head": III_HEAD["T303"], "msg": FIXED_MSG["T303"]},
    {"chap": "Ⅲ", "sys": False, "type": "T304", "ids": ["s23556", "s06141", "s21747", "s21985", "s21131"],
     "head": III_HEAD["T304"], "msg": FIXED_MSG["T304"]},
]
SYS_RE = r"talenx|SaaS|구독형|성과관리 ?(시스템|툴|솔루션|플랫폼)|시스템 (도입|구축|연계|셋업|세팅)|솔루션 (도입|연계)"


def intro_ids(cfg):
    v = str(cfg.get("intro", "그대로"))
    if "빼" in v or v in ("0", "none"):
        return []
    return [x for x, _ in INTRO[:4]] if ("4" in v or "줄" in v) else [x for x, _ in INTRO]


def weeks(info):
    """기간(주). 「24 주요 고객사」 같은 말은 빼고, 가장 많이 나온 값(같으면 큰 값)."""
    ws = [int(x) for x in re.findall(r"(\d{1,2})\s*주(?![요차식제기말도민체관력임장주])", info or "") if 2 <= int(x) <= 60]
    if not ws:
        return None
    cnt = collections.Counter(ws)
    return max(cnt, key=lambda w: (cnt[w], w))


def schedule_msg(info):
    w = weeks(info)
    return f"본 프로젝트는 다음과 같은 일정&절차에 따라 총 {w}주간 진행될 예정임" if w else "본 프로젝트는 다음과 같은 일정&절차에 따라 진행될 예정임"


def out(**k):
    print(json.dumps(k, ensure_ascii=False))


def jload(p, default=None):
    try:
        return json.load(open(p, encoding="utf-8-sig"))
    except (OSError, ValueError):
        return default


def jsave(p, o):
    json.dump(o, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


class Ctx:
    def __init__(self, data, work=None):
        self.data = os.path.abspath(data)
        self.cache = store.cache_dir(self.data)
        self.work = os.path.abspath(work) if work else None
        if self.work:
            os.makedirs(self.work, exist_ok=True)
            self.cfg = jload(self.p("client.json"), {}) or {}
            self.info = open(self.p("input.txt"), encoding="utf-8").read().strip() if os.path.exists(self.p("input.txt")) else ""

    def p(self, *a):
        return os.path.join(self.work, *a)

    def mods(self):
        """draft · post · merge 모듈을 이 작업에 맞춤."""
        import draft, post, merge
        draft.D = post.D = self.work
        draft.HERE = post.HERE = self.cache  # media_index.json
        draft.OUT = self.cache  # meta.json
        draft.CLIENT = post.CLIENT = self.cfg.get("short", "고객")
        draft.IND = self.cfg.get("ind", "")
        post.CLIENT_NAMES = self.cfg.get("names") or [self.cfg.get("short", "")]
        post.TITLE = self.cfg.get("title", "")
        post.INPUT = self.p("input.txt")
        post.META = os.path.join(self.cache, "meta.json")
        merge.LIB = os.path.join(self.cache, "pptx")
        return draft, post, merge


def log_step(c, name, t, **k):
    lg = jload(c.p("run_log.json"), {"steps": []})
    lg["steps"].append({"step": name, "sec": round(time.time() - t, 1), **k})
    jsave(c.p("run_log.json"), lg)


def ppt_busy():
    """사용자 PowerPoint 를 건드리지 않기 위한 확인. HCG_PPT_ASK=1 이면 떠 있을 때 멈춤(관리자 PC 시험용). 기본은 숨은 창으로 같이 씀(우리 파일만 닫고 끄지 않음)."""
    if os.environ.get("HCG_PPT_ASK") != "1":
        return False
    # 창이 있는 PowerPoint 만 사용자 것으로 봄. 검사용으로 띄운 숨은 PowerPoint(창 없음)는 끝난 뒤 몇 분 남기도 함
    cmd = "@(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }).Count"
    r = subprocess.run(["powershell", "-NoProfile", "-Command", cmd], capture_output=True, text=True, errors="ignore").stdout.strip()
    return r not in ("", "0")


def ps(script, *a, wait=True):
    args = PS + [os.path.join(S, script), *a]
    if wait:
        return subprocess.run(args, capture_output=True, text=True, encoding="utf-8", errors="ignore")
    return subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


# ---------------- setup ----------------
def setup(data):
    st = store.status(data)
    imported = None
    if not st.get("key") or not st.get("server"):
        imported = store.import_key(data)
        st = store.status(data)
    got = store.sync(data) if st.get("server") else []
    has_ppt = subprocess.run(["reg", "query", r"HKCR\PowerPoint.Application"], capture_output=True).returncode == 0
    out(ready=bool(st.get("server")), email=st.get("email"), why=st.get("why") or (imported or {}).get("why"), synced=got, powerpoint=has_ppt)


# ---------------- start ----------------
def words(s):
    return set(re.findall(r"[가-힣A-Za-z]{2,}", s or ""))


STOP = words("그리고 위한 대한 통해 있음 지금 기존 검토 따로 운영 희망 우려 주요 기간 분량 일반 목표 시행 고객사 회의 내용 스콥 약 명 제안서 프로젝트")
# 요청 글 → 과제 분류(창고 meta 의 mods). 기준 제안서 고를 때 과제가 맞는지 봄
TOPIC_RE = {"평가·성과관리": r"평가|성과관리|KPI|MBO|OKR|목표관리", "보상": r"보상|임금|연봉|성과급|인센티브|Pay|급여|수당",
            "직급·승진": r"직급|승진|호칭|직책", "조직문화": r"조직문화|기업문화|몰입|소통|심리적 안정",
            "조직 설계·진단": r"조직 ?진단|조직 ?구조|조직 ?설계|조직 ?개편|R&R", "HR 시스템·솔루션": r"e-HR|HRIS|인사 ?시스템 (구축|도입)|SaaS|talenx|솔루션 도입|시스템 구축",
            "직무": r"직무", "리더십·육성": r"리더십|육성|교육|경력개발", "인력·적정인력": r"적정 ?인력|인력 ?산정|인력 ?운영",
            "채용": r"채용", "노사·노무": r"노조|노무|노사|포괄임금"}


def req_topics(text):
    return {t for t, rx in TOPIC_RE.items() if re.search(rx, text or "", re.I)}


def start(data, work):
    c = Ctx(data, work)
    meta = jload(os.path.join(c.cache, "meta.json"), [])
    cfg = c.cfg
    excl = cfg.get("excl") or "|".join(map(re.escape, cfg.get("names") or [cfg.get("short", "@@")]))
    info = c.info
    if os.path.exists(c.p("docs.md")):  # 고객 자료를 먼저 읽었으면 그 글(RFP 등)도 봄
        info += "\n" + open(c.p("docs.md"), encoding="utf-8").read()[:20000]
    pw = words(info) - STOP
    head = (cfg.get("title") or "") + " " + (cfg.get("project") or "")
    want_t = req_topics(head)  # 과제 이름의 주제는 두 배로
    want = req_topics(info) | want_t
    docs = collections.defaultdict(list)
    for r in meta:
        if r.get("doc") and not re.search(excl, (r.get("client") or "") + (r.get("doc") or "")):
            docs[r["doc"]].append(r)
    ranks = []
    for doc, rs in docs.items():
        if len(rs) < 15 or not rs[0].get("client"):
            continue
        ind = collections.Counter(x.get("industry") for x in rs).most_common(1)[0][0]
        mods = collections.Counter(m for x in rs for m in (x.get("mods") or []))
        top = [m for m, _ in mods.most_common(3)]
        txt = " ".join((x.get("name") or "") + " " + (x.get("text") or "")[:200] for x in rs)
        # 낱말 겹침은 큰 덱일수록 커지므로 위를 막고, 과제 분류가 맞는 장 비율 · 안 맞는 주제(예: 시스템 구축 덱) 비율을 봄
        s = min(len(pw & words(txt)) * 0.15, 5) + (3 if ind == cfg.get("ind") else 0)
        if want:
            share = {m: n / len(rs) for m, n in mods.items()}
            s += 10 * sum(share.get(m, 0) * (2 if m in want_t else 1) for m in want) / max(1, len(want)) ** 0.5
            s -= 6 * sum(v for m, v in share.items() if m not in want and m != "공통" and v >= 0.3)
        s -= 2 if len(rs) > 120 else 0
        s -= 0 if re.search(r"제안|proposal", doc, re.I) else 8  # 제안서가 아닌 자료(소개 자료 · 산출물 등)는 기준으로 쓰지 않음
        try:
            s += max(0, int(rs[0].get("year") or 0) - 2018) * 0.5
        except ValueError:
            pass
        why = []
        if ind == cfg.get("ind"):
            why.append(f"같은 업종({ind})")
        hit = [m for m in top if any(w in m for w in pw) or any(m.split("·")[0] in w for w in pw)]
        if hit:
            why.append("같은 과제(" + " · ".join(hit[:2]) + ")")
            s += 2 * len(hit)
        main_t = [m for m in top[:2] if m in want_t]  # 과제 이름의 주제가 그 덱의 주요 주제
        if main_t:
            s += 4
            if not hit:
                why.append("같은 과제(" + main_t[0] + ")")
        elif want_t:
            s -= 3  # 과제 이름의 주제(예: 조직진단)가 그 덱의 주요 주제가 아님
        ranks.append((s, {"doc": doc, "client": rs[0]["client"], "year": rs[0].get("year"), "slides": len(rs), "industry": ind, "mods": top, "why": " · ".join(why) or "내용이 가까움"}))
    ranks.sort(key=lambda x: -x[0])
    inds = sorted({r.get("industry") for r in meta if r.get("industry")})
    listed = None
    if cfg.get("dart"):
        try:
            import dart
            dart.get(dart.BASE + "/dsab001/main.do")
            import urllib.parse
            q = urllib.parse.urlencode({"textCrpNm": cfg["dart"], "currentPage": 1, "maxResults": 15, "maxLinks": 10, "sort": "date", "series": "desc",
                                        "startDate": time.strftime("%Y%m%d", time.localtime(time.time() - 400 * 86400)), "endDate": time.strftime("%Y%m%d"),
                                        "publicType": "A001", "finalReport": "recent"}).encode()
            h = dart.get(dart.BASE + "/dsab001/search.ax", q, dart.BASE + "/dsab001/main.do")
            listed = "사업보고서" in h and cfg["dart"].replace(" ", "") in re.sub(r"\s", "", h)
        except Exception:
            listed = None
    out(proposals=[r for _, r in ranks[:3]], listed=listed, industries=inds)


# ---------------- docs ----------------
def docs_cmd(data, work, folder=""):
    """고객 자료 폴더의 파일 글을 뽑아 docs.md 로(원본 폴더는 읽기만). 폴더는 인자 또는 client.json 의 docs."""
    c = Ctx(data, work)
    folder = (folder or c.cfg.get("docs") or "").strip().strip('"')
    if not folder or not os.path.isdir(folder):
        out(error=f"자료 폴더를 찾지 못함: {folder or '(비어 있음)'}"); return
    import docs
    t = time.time()
    inside = os.path.abspath(folder).lower().startswith(c.work.lower())
    rows, total = docs.main(folder, c.p("docs.md"), None if inside else c.work)
    c.cfg["docs"] = folder
    jsave(c.p("client.json"), c.cfg)
    log_step(c, "자료 읽기", t, files=len(rows), chars=total)
    out(file=c.p("docs.md"), files=len(rows), chars=total,
        read=[f"{r['file']} ({r['chars']:,}자)" for r in rows if r.get("chars")],
        failed=[f"{r['file']} — {r['error']}" for r in rows if r.get("error")])


# ---------------- plan ----------------
def excl_re(cfg):
    return cfg.get("excl") or "|".join(map(re.escape, cfg.get("names") or [cfg.get("short", "@@")]))


def std_rows(c, meta=None):
    """HCG 표준 장(Why HCG · 수행 조직 · 투입 인력) — 이번 고객 제안서 것은 건너뛰고 앞 id 부터."""
    if meta is None:
        meta = jload(os.path.join(c.cache, "meta.json"), [])
    by = {r["id"]: r for r in meta}
    ex, sys_ = excl_re(c.cfg), bool(re.search(SYS_RE, c.info, re.I))
    rows = []
    for w in STD_ROWS:
        if w["sys"] and not sys_:
            continue
        for sid in w["ids"]:
            r = by.get(sid)
            if r and not re.search(ex, (r.get("client") or "") + (r.get("doc") or "")):
                rows.append({"chap": w["chap"], "type": w["type"], "pick": sid, "head": w["head"], "msg": w["msg"], "why": "HCG 표준 장", "std": True})
                break
    return rows


def chapters(cfg, nstd=0, niii=2):
    """분량 = Ⅰ~Ⅲ장 본문 장수(HCG 소개 · 목차 · 간지 빼고). 계획할 장수 i(Ⅰ) · ii(Ⅱ) — 표준 장 nstd · Ⅲ장 표준 niii 는 빼고."""
    n = int(cfg.get("body") or 40)
    iv = len(intro_ids(cfg))
    rest = max(8, n - 1 - nstd - niii)  # 1 = 추진 일정(T301)
    i = max(4, round(rest * 0.28))
    return n, i, rest - i, 1 + niii, iv


def cands(c, meta=None, skip=()):
    if meta is None:
        meta = jload(os.path.join(c.cache, "meta.json"), [])
    types = jload(os.path.join(c.cache, "types.json"), {}).get("types", [])
    cfg = c.cfg
    excl = excl_re(cfg)
    pw = words(c.info) - STOP
    keep = [t for t in types if t["ch"] in ("1", "2", "3", "4") and t["code"] not in ("T302", "T303", "T304", "T401", "T402", "T403", "T404")]  # 회사 소개 · 수행 조직 · 투입 인력 · 관리방안은 표준 장
    codes = {t["code"] for t in keep}
    by = collections.defaultdict(list)
    for r in meta:
        if r.get("type") not in codes or r["id"] in skip or re.search(excl, (r.get("client") or "") + (r.get("doc") or "")):
            continue
        s = len(pw & words((r.get("name") or "") + " " + (r.get("text") or "")[:600])) * 1.0
        s += 3 if r.get("std") else 0
        s += 2 if r.get("star") else 0
        s += {"상": 2, "중": 0.5}.get(r.get("grade"), 0)
        s += 2 if r.get("industry") == cfg.get("ind") else 0
        s -= min(int(r.get("opic") or 0), 5) * 0.8  # 다른 고객 그림이 많으면 사람이 바꿔 넣을 일이 많음
        s += 4 if cfg.get("base_doc") and r.get("doc") == cfg["base_doc"] else 0
        try:
            s += max(0, int(r.get("year") or 0) - 2019) * 0.4
        except ValueError:
            pass
        s += 1.5 if r.get("tconf") == "H" else 0
        by[r["type"]].append((s, r))
    lines = []
    diag = bool(re.search(r"진단", (cfg.get("title") or "") + " " + (cfg.get("project") or "")))
    for t in keep:
        rows = sorted(by.get(t["code"], []), key=lambda x: -x[0])
        k = 12 if diag and t["code"] in ("T202", "T204", "T205", "T206", "T207", "T208") else 10 if t["code"] in ("T211", "T216", "T105", "T201", "T301", "T407") else (8 if t["code"] in ("T101", "T102", "T104", "T204", "T205", "T206", "T207", "T405", "T406", "T219", "T220") else 6)
        seen, picked = set(), []
        for s, r in rows:
            if r.get("fam") and (r["fam"], r.get("var")) in seen:
                continue
            seen.add((r.get("fam"), r.get("var"))); picked.append(r)
            if len(picked) >= k:
                break
        if picked:
            lines.append(f"## {t['code']} {t['name']} — {t['desc']}")
            for r in picked:
                tx = re.sub(r"\s+", " ", r.get("text") or "")[:110]
                op = int(r.get("opic") or 0)
                lines.append(f"{r['id']}|{r.get('name') or r.get('title')}|{r.get('client')} {r.get('year')} {r.get('industry')}{f' 그림{op}' if op else ''}|{tx}")
    return "\n".join(lines)


PLAN_RULES = """# 규칙
- 장 수: Ⅰ 제안 개요 {i}장 → Ⅱ 수행 방안 {ii}장 → Ⅲ 추진 일정 1장. 표지 · 목차 · 간지는 계획하지 않는다.
  저절로 붙는 표준 장(계획에 넣지 않음): {std}.{fixed}{base}
- Ⅰ장 흐름(이 순서, 장마다 다른 이야기):
  ① 고객 이해 · 추진 배경 1~2장(T101 · T103) — 고객 사업 현황 · 업황 · 공시 숫자에서 과제 필요성으로. 공시 숫자가 있으면 매출 · 이익 추이 차트가 있는 장 우선.
  ② 환경 변화 · HR Trend 1장(T102) — 과제 주제(평가 · 보상 · 직급 · 조직문화 등)의 최근 흐름. 장이 넉넉하면 [참고] 국내 선도기업 HR Trend 1장 더.
  ③ 現 제도 이슈 · 가설적 이슈 1~2장(T104 · T105 · T210) — As-is/To-be 대비 장 우선. 고객 내부 방향이 정해지기 전이면 To-be 는 「(잠정 案)」.
  ④ 프로젝트 목적 및 범위 1장(T106) — 범위 = Ⅱ장 모듈 순서 그대로.
  ⑤ Why HCG(이 순서): 「제안 특징 및 장점    Why HCG」 요약 1장(T406 — 강점 3가지) → 「Why HCG    1. 주제에 대한 전문성 …」 1장(T405) → [참고] 유사 프로젝트 수행 사례 1장(T405 · T206) → 「Why HCG    2. …」 — 범위에 시스템(talenx · SaaS)이 없으면 2번 강점 1장을 계획한다(T407 Cutting-Edge 방법론 · Data Analytics, 또는 Benchmarking 전문성 · 고객사 이해). 시스템이 있으면 2번은 표준 장(자체 솔루션)이 붙는다. 3번 「지속적인 변화관리 지원」은 표준 장이 붙는다.
  배경 = 회사 상황, 트렌드 = 바깥 흐름, 이슈 = 제도 문제, 목적 = 해결 방향. 같은 이슈를 여러 장에서 되풀이하지 않는다. 추진 전략(T108) · 기대효과(T109)는 장이 넉넉할 때만.
- Ⅱ장 흐름(모듈마다 번호):
  ① (장이 넉넉하면) Overall Approach 1장(T201) — 단계 = 아래 모듈 순서.
  ② 「1. 현황 진단 및 방향성 수립」 모듈 — 기본으로 넣는다: Approach(Document Review · 인터뷰/서베이 · 벤치마킹, T204) · Interview & Survey(T205) · Benchmarking(T206) · 이슈 및 개선 방향성 도출(T207 · T208) 중 2~4장. 고객이 진단을 최소로 해 달라고 했을 때만 「現 제도 분석」 1장으로 줄인다.
     과제 이름이 진단(조직진단 · 조직문화 진단 · 인사 진단 등)이면 진단 모듈이 Ⅱ장의 중심이다: Ⅱ장의 40~50%를 진단 방향 · 진단 Framework(1~2장) · 진단 항목별 Survey/Interview 설계(2~4장) · 조직운영 · 데이터 분석 · Benchmarking · 진단 결과 이미지 · 개선 방향성으로 쓰고, 뒤 개선안 설계 모듈은 모듈마다 3~4장.
  ③ 「2. …」 「3. …」 = 요청 범위 항목을 고객 RFP · 회의 순서대로 모듈로. 모듈마다 2~6장 — 변화 Trend → 설계 방향 · Framework → 상세 설계(유형 · 대안 · 기준) → Simulation · 연계 방안 순. 주요 모듈 설계 장 뒤에 [참고] 장(T407 HCG 자산 · T405 · T206 수행 사례 · T220 산출물 예시). [참고]는 Ⅱ장의 15~25%.
  ④ 마지막 모듈 「N. 변화관리 및 실행지원」 1~2장(T219 — 단계별 Roll-out · 커뮤니케이션, 시스템 범위면 시스템 셋업).
- 장마다 「후보」에서 장 하나(id)를 고른다. 같은 id 두 번 금지. 유형이 맞고 내용이 가까운 장(업종 · 스콥 · 회의 내용). 도식(그림 속 라벨)이 이번 메시지와 다른 논리인 장(예: 이슈는 평가 목적 혼재인데 도식은 등급 분포)은 고르지 않는다. 다른 고객 이야기가 깊게 박힌 장보다 고치기 쉬운 장. 후보 끝의 「그림N」은 사람이 바꿔 넣어야 할 다른 고객 그림 수 — 내용이 비슷하면 그림이 적은(없는) 장.
- 어깨끈(head) — 장 성격에 따라, 나누는 곳은 공백 4칸. 고객사명은 넣지 않는다:
  · Ⅰ장 · Overall Approach · 추진 일정 = 「제안 배경    장 이름」 또는 장 이름 하나(예: 「제안 배경    사업현황」 「제안 배경    HR Trend」 「프로젝트 목적 및 범위」 「Overall Approach」 「추진 일정」)
  · Why HCG 장 = 「제안 특징 및 장점    Why HCG」 · 「Why HCG    1. 주제에 대한 전문성 보유」 · 「Why HCG    2. Cutting-Edge 방법론」처럼 번호
  · Ⅱ장 설계 장 = 「N. 모듈    장 이름」 또는 「N. 모듈    세부    장 이름」(예: 「1. 현황 진단 및 방향성 수립    Approach」 「3. 보상제도 개선    기본급 인상    평가인상 방식」). 번호 · 모듈 이름은 목적 및 범위 · Overall Approach 와 같게.
  · [참고] 장 = 「[참고] 장 이름」 하나(예: 「[참고] HCG 역량모델 활용」 「[참고] 유사프로젝트 수행사례    D사」)
- 핵심 메시지(msg) = 2줄, 70~110자. 이유(고객 상황 · 이슈 · 요청)에서 결론(할 일 · 효과)으로 잇는다. 예: 「반도체 업황 회복으로 성장세가 이어지고 있어, 중장기 성장 목표의 실행력을 높이기 위한 성과관리 체계 고도화가 필요한 시점」 「업무특성과 수행방식에 따른 성과 창출 방식의 차이를 고려하여 사무/생산의 평가방식을 구체화 함」. 끝은 「~함 · ~임 · ~시점」. 고객 정보 · 고객 자료의 사실 · 숫자만 쓰고 지어내지 않는다.
- 추진 일정(T301) 핵심 메시지는 통상 문구 그대로: 「{sched}」.
- why = 그 장을 고른 까닭 20자 안.
- facts = 고객 정보 · 고객 자료에서 장표에 쓸 사실 · 숫자 8~20개(한 줄씩, 끝에 출처 괄호). 현행 제도(평가 항목 · 반영 비율 · 등급 배분 · 시기), 인원, 경영 숫자, 일정, 고객 우려 위주. 자료에 없는 것은 쓰지 않는다.
"""


def plan(data, work):
    c = Ctx(data, work)
    t = time.time()
    meta = jload(os.path.join(c.cache, "meta.json"), [])
    std = std_rows(c, meta)
    jsave(c.p("std.json"), std)
    nstd1 = len([r for r in std if r["chap"] == "Ⅰ"])
    n, i, ii, iii, iv = chapters(c.cfg, nstd1, len(std) - nstd1)
    fixed = f"\n  Ⅳ HCG 소개({iv}장)는 정해진 표준 장을 그대로 넣으므로 계획하지 않는다." if iv else "\n  HCG 소개 장은 넣지 않는다."
    base = f"\n- 기준 제안서: {c.cfg['base_doc']} — 이 제안서의 장 흐름을 참고하고, 후보에 있으면 그 장을 먼저 고려한다(위 흐름 규칙이 먼저)." if c.cfg.get("base_doc") else ""
    std_txt = " · ".join(f"Ⅰ장 끝 「{r['head'].split('    ')[-1]}」" if r["chap"] == "Ⅰ" else f"Ⅲ장 「{r['head']}」" for r in std) or "없음"
    docs_txt = ""
    if os.path.exists(c.p("docs.md")):
        docs_txt = "\n\n# 고객 자료(파일에서 뽑은 글 · 앞부분)\n" + open(c.p("docs.md"), encoding="utf-8").read()[:24000]
    rules = PLAN_RULES.format(i=i, ii=ii, std=std_txt, fixed=fixed, base=base, sched=schedule_msg(c.info))
    prompt = (f"너는 HCG(인사 컨설팅사) 제안서 Draft 를 짜는 시니어 컨설턴트다. 아래 고객 정보로 Ⅰ~Ⅲ장 본문 {i + ii + 1}장 계획을 만든다"
              f"(표준 장 {len(std)}장은 저절로 붙어 모두 {i + ii + 1 + len(std)}장).\n\n# 고객 정보\n{c.info}{docs_txt}\n\n{rules}\n"
              f"# 후보(형식: id|장 이름|고객 연도 업종 [그림N = 바꿔 넣어야 할 다른 고객 그림 N개]|본문 앞부분)\n{cands(c, meta, {r['pick'] for r in std})}\n\n# 출력\n"
              'plan.json 에 JSON 객체 하나만 쓴다: {"facts": ["현행 평가 = 업적 · 직무수행능력 · 다면 종합(통화회의록)", ...], '
              '"slides": [{"no":1,"chap":"Ⅰ","type":"T101","pick":"s01888","head":"프로젝트 추진배경","msg":"...","why":"..."}, ...]}')
    open(c.p("plan_prompt.md"), "w", encoding="utf-8").write(prompt)
    disc = {"items": []}
    if c.cfg.get("numbers", "dart") == "dart" and c.cfg.get("dart"):
        import dart
        try:
            dart.main(c.cfg["dart"], c.p("disc.json"))
            disc = jload(c.p("disc.json"), disc)
        except Exception as e:
            disc = {"items": [], "note": f"DART 를 못 읽음: {e}"}
            jsave(c.p("disc.json"), disc)
    else:
        jsave(c.p("disc.json"), disc)
    log_step(c, "후보 · 공시", t)
    out(prompt=c.p("plan_prompt.md"), plan=c.p("plan.json"), disc=[f"{x['key']} {x.get('year') or ''} {x['value']}" for x in disc.get("items", [])], disc_note=disc.get("note") or disc.get("report"))


def normalize_plan(c):
    """계획 다듬기: {facts, slides} → facts.json + 장 목록, Ⅳ장(고정)은 빼고, 표준 장(Why HCG → Ⅰ장 끝 · 수행 조직 · 투입 인력 → Ⅲ장)을 붙이고,
    Ⅲ장 어깨끈 · 핵심 메시지는 통상 문구로, 장 번호(no)는 처음부터 다시. plan.json 을 장 목록으로 고쳐 씀(여러 번 불러도 같음)."""
    raw = jload(c.p("plan.json"), []) or []
    if isinstance(raw, dict):
        jsave(c.p("facts.json"), [str(x) for x in (raw.get("facts") or [])])
        raw = raw.get("slides") or []
    std = jload(c.p("std.json"), []) or []
    std_ids = {r["pick"] for r in std}
    pl = [p for p in raw if p.get("chap") in ("Ⅰ", "Ⅱ", "Ⅲ") and (p.get("std") or p.get("pick") not in std_ids)
          and (p.get("std") or p.get("type") not in ("T302", "T303", "T304"))]
    if std and not any(p.get("std") for p in pl):
        k = max((x for x, p in enumerate(pl) if p.get("chap") == "Ⅰ"), default=-1)
        pl = pl[:k + 1] + [dict(r) for r in std if r["chap"] == "Ⅰ"] + pl[k + 1:]
        pl += [dict(r) for r in std if r["chap"] == "Ⅲ"]
    order = {"Ⅰ": 0, "Ⅱ": 1, "Ⅲ": 2}
    pl.sort(key=lambda p: order[p["chap"]])  # 안정 정렬 — 장 안 순서는 그대로
    for p in pl:
        t = p.get("type")
        if p["chap"] == "Ⅲ" and t in III_HEAD:
            p["head"] = III_HEAD[t]
        if t in FIXED_MSG:
            p["msg"] = FIXED_MSG[t]
        elif t == "T301" and not re.search(r"진행될 예정임$", str(p.get("msg", ""))):
            p["msg"] = schedule_msg(c.info)
    for k, p in enumerate(pl, 1):
        p["no"] = k
    jsave(c.p("plan.json"), pl)
    return pl


MSG_MIN, MSG_MAX = 70, 115


def plan_check(data, work):
    """계획의 핵심 메시지 길이 확인. Ⅰ · Ⅱ장(표준 장 빼고) 메시지가 70자 미만 · 115자 넘는 장이 4장 이상이면
    그 장만 고쳐 쓰게 하는 요청문(plan_fix.md)을 만듦 — hcg-planner 가 plan.json 을 고쳐 씀."""
    c = Ctx(data, work)
    pl = normalize_plan(c)
    bad = [p for p in pl if not p.get("std") and p.get("chap") in ("Ⅰ", "Ⅱ") and not (MSG_MIN <= len(str(p.get("msg") or "")) <= MSG_MAX)]
    if len(bad) < 4:
        out(bad=len(bad), prompt=None); return
    rows = "\n".join(f"- no {p['no']} · {p['head']} · 지금 {len(str(p.get('msg') or ''))}자: {p.get('msg')}" for p in bad)
    facts = "\n".join(f"- {x}" for x in jload(c.p("facts.json"), [])[:20])
    prompt = (f"# 할 일\n{c.p('plan.json')} 은 제안서 Draft 계획(JSON 배열)이다. 아래 장들의 핵심 메시지(msg)만 고쳐 쓰고, 같은 파일에 배열 그대로 다시 쓴다. "
              "다른 장 · 다른 칸(no · chap · type · pick · head · why · std)은 한 글자도 바꾸지 않는다.\n\n"
              "# msg 규칙\n- 두 줄, 띄어쓰기 포함 75~105자. 이유(고객 상황 · 이슈 · 요청)에서 결론(할 일 · 효과)으로 잇는다. 끝은 「~함 · ~임 · ~시점」.\n"
              "- 그 장의 어깨끈(head) 내용에 맞게. 고객 정보 · 찾은 사실에 있는 것만 쓰고 지어내지 않는다.\n"
              "- 예: 「반도체 업황 회복으로 성장세가 이어지고 있어, 중장기 성장 목표의 실행력을 높이기 위한 성과관리 체계 고도화가 필요한 시점」(75자)\n\n"
              f"# 고칠 장\n{rows}\n\n# 찾은 사실\n{facts}\n\n# 고객 정보\n{c.info[:6000]}\n")
    open(c.p("plan_fix.md"), "w", encoding="utf-8").write(prompt)
    out(bad=len(bad), prompt=c.p("plan_fix.md"), plan=c.p("plan.json"))


def plan_table(data, work):
    c = Ctx(data, work)
    meta = {r["id"]: r for r in jload(os.path.join(c.cache, "meta.json"), [])}
    pl = normalize_plan(c)
    rows = ["| 장 | 어깨끈 | 핵심 메시지 | 가져올 장 |", "|---|---|---|---|"]
    for p in pl:
        m = meta.get(p.get("pick"), {})
        src = "HCG 표준" if p.get("std") else f"{m.get('client') or '?'} {m.get('year') or ''}"
        rows.append(f"| {p.get('chap')}-{p.get('no')} | {' · '.join(str(p.get('head', '')).split('    '))} | {p.get('msg', '')} | {src} |")
    ids = intro_ids(c.cfg)
    if ids:
        names = dict(INTRO)
        rows.append(f"| Ⅳ | HCG 소개 | {' · '.join(names[x] for x in ids)} — 표준 장 그대로 | HCG 표준 |")
    print("\n".join(rows))


# ---------------- build ----------------
RW_RULES = """# 할 일
아래 장들은 다른 고객 제안서에서 가져온 장이다. 글을 「{short} {title}」 제안서 Draft 로 고쳐 쓴다. 사람이 조금만 손보면 낼 수 있는 수준.

# 규칙
1. 장 구조 · 항목 수는 그대로. **원문보다 길게 쓰지 않는다** — 항목마다 원문 길이의 70~110% 안(칸 넘침 막기). 상자 · 원 안의 짧은 문구는 원문보다 짧게.
2. 원문 글에 줄바꿈(\\n)이 있으면 같은 줄 수로 쓰고 줄마다 원문 그 줄 길이를 넘지 않게. 원문에 줄바꿈이 없으면 넣지 않음.
3. 어깨끈 · 핵심 메시지(장 맨 위 제목 자리)는 계획대로 따로 넣으므로 고칠 항목에 없다. 고칠 항목만 쓴다.
4. 다른 고객 이름 · 사업 · 제도 이름 · 수치 · 연혁 → 이번 고객 정보로. 숫자는 「고객 정보」와 「공시 숫자」에 있는 값만 실제 숫자로 쓰고, 없는 사실(세부 인원 · 연혁 · 지표 등)은 지어내지 말고 「○○」. 회의 메모와 공시 숫자가 다르면 공시 숫자를 쓰고 checks 에 차이를 적음.
5. 방법론 · 절차 · HCG 자산은 내용을 살리고 다른 고객 언급만 고침. 사례 · 벤치마킹으로 남길 회사 이름은 「A社」처럼 익명.
6. 목차 · 간지: 장 이름 = Ⅰ 제안 개요 · Ⅱ 수행 방안 · Ⅲ 수행 체계{iv_name}. 간지 Ⅲ 의 작은 목록 = 추진 일정 · 수행 조직 · 투입 인력(남는 칸은 ""). 표지 = 「{project}」 · 「- 제안서 -」 · 「{date}」.
7. 차트: 공시 숫자로 채울 수 있는 차트(매출 · 영업이익 · 인원 추이 등)면 그 값과 연도로 바꾸고 제목 끝에 「(DART)」. 공시 숫자로 못 채우거나 이 차트 주제가 고객에 맞지 않으면 values 를 모두 null(빈 값)로 두고 지어내지 않음 — checks 에 「차트 맞는지 · 값 넣기」.
8. 지울 항목은 "".
9. 굵게: 글 안 **…** 는 원문에서 굵게(또는 다른 서식)였던 구절이다. 새 글에서도 같은 성격의 구절을 **…** 로 감싼다(개수는 원문과 비슷하게, 원문에 없으면 넣지 않음). 다만 원 · 상자 안 2~3줄 짧은 문구에는 원문에 없어도 결론 구절 하나를 **…** 로 굵게 한다.
10. 사람이 확인해야 할 곳(공시와 회의 메모가 다른 숫자 · 사실 확인 · 고객 진단 값 · 다른 회사 예시가 남은 곳)은 checks 에 쓴다. 장 위에 노란 딱지로 붙으니 **자리(id)와 짧은 딱지 말(tag, 10자 안)**을 꼭 넣는다. ○○ 자리는 저절로 딱지가 붙으니 적지 않아도 된다. checks 는 장마다 꼭 필요한 것만(장당 2개 안).
11. 「HCG 표준 장」 · Why HCG · [참고] 장(HCG 실적 · 솔루션 · 방법론 · 실행지원 · 수행 사례)은 HCG 자기 사실이다. 본문 내용 · 숫자는 그대로 두고, 다른 고객 이름만 「A社」식 익명으로, 원래 제안서 번호(「Why HCG 3.」 「HCG 특장점 2.」)는 지운다. 사례 장은 끝 문장에 이번 고객 적용 문구를 한 줄 넣을 수 있다.
12. 도식 라벨(도형 · 원 · 화살표 안 짧은 말 · 표 머리)도 항목에 있으면 이번 고객 논리에 맞춘다. 원래 도식이 이번 핵심 메시지와 다른 논리(예: 이슈는 목적 혼재인데 도식은 등급 분포)라 라벨만으로 못 맞추면, 라벨을 이번 이슈 말로 바꾸고 checks 에 「도식 바꾸기」.
13. 「고객 자료에서 찾은 사실」의 현행 제도 · 인원 · 일정 숫자는 그대로 써도 된다(출처는 각주 자리에 「1) 고객 자료(파일 이름)」).
14. Overall Approach · 프로젝트 목적 및 범위 · 추진 일정 · 이행계획 장: 단계 이름 · 과업은 「덱 전체 흐름」의 Ⅱ장 모듈(번호 · 이름) 순서와 「고객 정보」의 요청 범위대로 쓴다. 원문 장에 있던 이번 계획에 없는 모듈 과업(예: 계획에 없는 직무분석 · 조직구조 개선 · 시스템 구축)은 남기지 말고 그 칸을 이번 모듈 과업(예: talenx 설정 · 노조 브리핑 · 직원 설명회)으로 채운다. 현황 진단(Document Review · 인터뷰/서베이 · 벤치마킹)은 덱 흐름에 진단 모듈이 있을 때만 남긴다. 칸이 남으면 "".
15. 투입 인력 장: 이름 · 학력 · 경력 · 주요 고객 칸은 「○○」로 비우고, Position 열만 「고객 정보」의 투입 인력(예: PD 1 · PM 1 · 컨설턴트 2)에 맞춘다. 줄이 남으면 "" 로 비운다. 수행 조직 장도 사람 이름은 「○○○」.

# 출력
지정한 답 파일에 JSON 하나만 쓴다: {{"texts": {{"항목id": "새 글", ...}}, "charts": {{"차트id": {{"title": "...", "categories": [...], "series": [{{"name": "...", "values": [...]}}]}}}}, "checks": [{{"slide": 4, "id": "4:12:0", "tag": "인원 수 확인", "note": "..."}}]}}
바꾼 항목만 넣음(그대로 둘 항목은 빼도 됨). 글 안 줄바꿈은 JSON 문자열의 \\n 으로.
"""


def mod_no(head):
    """「3. 보상제도 개선    기본급」 → (3, "보상제도 개선"). 번호 없는 어깨끈([참고] · Overall Approach)은 None."""
    m = re.match(r"\s*(\d+)\.\s*(\S.*?)(?:\s{2,}|$)", str(head or ""))
    return (int(m.group(1)), m.group(2).strip()) if m else None


def modules(picks):
    """Ⅱ장 계획의 모듈 (번호, 이름) — 처음 나온 순서."""
    res, seen = [], set()
    for p in picks:
        k = mod_no(p.get("head"))
        if k and k[0] not in seen:
            seen.add(k[0]); res.append(k)
    return res


def add_modlists(deck, order, frame):
    """모듈 간지(mods 가 있는 간지 Ⅱ)에 모듈 목록 상자를 넣음 — 틀의 간지 Ⅲ 작은 목록 상자를 본떠, 지금 모듈만 굵게.
    상자 이름 HCG_MODS(고쳐 쓰기 · 서식 맞추기에서 뺌)."""
    if not any(o.get("mods") for o in order):
        return
    import copy
    from pptx import Presentation
    fr = Presentation(frame)
    src = next((sh for sh in fr.slides[FRAME["Ⅲ"] - 1].shapes if sh.has_text_frame and len(sh.text_frame.paragraphs) >= 3), None)
    if src is None:
        return
    prs = Presentation(deck)
    for s, o in zip(prs.slides, order):
        if not o.get("mods"):
            continue
        el = copy.deepcopy(src._element)
        ids = [int(x) for x in s.shapes._spTree.xpath(".//p:cNvPr/@id")] or [1]
        el.xpath("./p:nvSpPr/p:cNvPr")[0].set("id", str(max(ids) + 1))
        el.xpath("./p:nvSpPr/p:cNvPr")[0].set("name", "HCG_MODS")
        s.shapes._spTree.append(el)
        body = el.xpath("./p:txBody")[0]
        paras = el.xpath("./p:txBody/a:p")
        on, off = paras[0], paras[1]
        for p in paras:
            body.remove(p)
        for k, name in enumerate(o["mods"]):
            p = copy.deepcopy(on if k == o["cur"] else off)
            runs = p.xpath("./a:r")
            for r in runs[1:]:
                p.remove(r)
            for x in p.xpath("./a:fld | ./a:br"):
                p.remove(x)
            runs[0].xpath("./a:t")[0].text = name
            rpr = runs[0].xpath("./a:rPr")
            if rpr:
                rpr[0].set("b", "1" if k == o["cur"] else "0")
            body.append(p)
    prs.save(deck)


def build(data, work):
    c = Ctx(data, work)
    draft, post, merge = c.mods()
    t = time.time()
    pl = normalize_plan(c)
    n, i, ii, iii, iv = chapters(c.cfg)
    order, used = [{"base": FRAME["표지"], "role": "표지"}], set()
    for ch in ("Ⅰ", "Ⅱ", "Ⅲ"):
        picks = [p for p in pl if p.get("chap") == ch and re.fullmatch(r"s\d+", str(p.get("pick", ""))) and p["pick"] not in used]
        if not picks:
            continue
        mods = modules(picks) if ch == "Ⅱ" else []
        names = [m for _, m in mods]
        first = {"base": FRAME[ch], "role": f"간지 {ch}"}
        if len(mods) >= 2:  # Ⅱ장 간지 = 모듈 목록(1번 강조), 2번 모듈부터는 모듈마다 간지 한 장 더
            first.update(mods=names, cur=0)
        order.append(first)
        seen = {mods[0][0]} if mods else set()
        for p in picks:
            try:
                store.slide(c.data, p["pick"])
            except Exception:
                continue
            k = mod_no(p.get("head"))
            if len(mods) >= 2 and k and k[0] not in seen:
                seen.add(k[0])
                order.append({"base": FRAME["Ⅱ"], "role": "간지 Ⅱ", "mods": names, "cur": [n for n, _ in mods].index(k[0])})
            used.add(p["pick"]); order.append({"lib": p["pick"], "role": "본문", "no": p["no"]})
    intro = []
    for sid in intro_ids(c.cfg):  # HCG 소개 — 표준 장 그대로
        try:
            store.slide(c.data, sid); intro.append(sid)
        except Exception:
            continue
    if intro:
        order.append({"base": FRAME["Ⅳ"], "role": "간지 Ⅳ"})
        order += [{"lib": sid, "role": "소개"} for sid in intro]
    order.append({"base": FRAME["끝장"], "role": "끝장"})
    jsave(c.p("order.json"), order)
    merge.main(c.p("order.json"), c.p("merged.pptx"), os.path.join(c.cache, "frame_std.pptx"))
    add_modlists(c.p("merged.pptx"), order, os.path.join(c.cache, "frame_std.pptx"))
    draft.extract(c.p("merged.pptx"), c.p("order.json"), c.p("plan.json"))
    slides = jload(c.p("items.json"), [])
    if ppt_busy():
        out(error="PowerPoint 가 열려 있어 멈춤"); return
    ps("boxes.ps1", "-File", c.p("merged.pptx"), "-Out", c.p("box_base.json"), wait=False)  # 원본 자리 재기 — 고쳐 쓰는 동안 뒤에서
    disc = jload(c.p("disc.json"), {}) or {}
    if disc.get("items"):
        disc_txt = "\n".join(f"- {x['key']} {x.get('year') or ''}: {x['value']} {x.get('unit') or ''}".strip() for x in disc["items"]) + f"\n(DART 원문 · {disc.get('report', '')})"
    else:
        disc_txt = "(없음 — 숫자 칸은 ○○, 차트는 빈 값)"
    facts = jload(c.p("facts.json"), []) or []
    facts_txt = "\n".join(f"- {x}" for x in facts) if facts else "(없음)"
    std_no = {p["no"] for p in pl if p.get("std")}
    outline =[f"{s['slide']} · {s['plan']['head']} | {s['plan']['msg']}" if "plan" in s else f"{s['slide']} · {s['role']}" + (" (HCG 표준 장 그대로)" if s["role"] == "소개" else "") for s in slides]
    rules = RW_RULES.format(short=c.cfg.get("short"), title=c.cfg.get("title"), project=c.cfg.get("project"), date=c.cfg.get("date"),
                            iv_name=" · Ⅳ HCG 소개" if iv else "")
    work_slides = [s for s in slides if s["items"] or s["charts"]]  # 그대로 넣는 장은 고쳐 쓰지 않음
    total = sum(sum(len(x["t"]) for x in s["items"]) + 200 for s in work_slides)
    groups, cur, acc = [], [], 0
    for s in work_slides:
        w = sum(len(x["t"]) for x in s["items"]) + 200
        if cur and acc + w > total / 5 * 1.05 and len(groups) < 4:
            groups.append(cur); cur, acc = [], 0
        cur.append(s); acc += w
    groups.append(cur)
    prompts = []
    for g, grp in enumerate(groups, 1):
        body = []
        for s in grp:
            hd = f"### 장 {s['slide']} · {s['role']}"
            if "plan" in s:
                hd += f" · 원본 {s['from']}{' · HCG 표준 장(규칙 11)' if s.get('no') in std_no else ''} · 계획 어깨끈「{s['plan']['head']}」 · 핵심 메시지「{s['plan']['msg']}」"
            body.append(hd)
            body += [f"{x['id']} (y{x['y']}) | {json.dumps(x['t'], ensure_ascii=False)}" for x in s["items"]]
            body += [f"차트 {ch['id']} | 제목「{ch['title']}」 | 범주 {json.dumps(ch['categories'], ensure_ascii=False)} | 계열 {json.dumps(ch['series'], ensure_ascii=False)}" for ch in s["charts"]]
        txt = (f"너는 HCG(인사 컨설팅사) 제안서를 쓰는 컨설턴트다.\n\n# 고객 정보\n{c.info}\n\n# 고객 자료에서 찾은 사실\n{facts_txt}\n\n# 공시 숫자\n{disc_txt}\n\n# 덱 전체 흐름({len(slides)}장)\n"
               + "\n".join(outline) + "\n\n" + rules + f"\n# 답 파일\n{c.p(f'ans_{g}.json')}\n\n# 고칠 장\n" + "\n".join(body))
        open(c.p(f"rw_{g}.md"), "w", encoding="utf-8").write(txt)
        prompts.append({"prompt": c.p(f"rw_{g}.md"), "answer": c.p(f"ans_{g}.json")})
    log_step(c, "합치기 · 요청문", t, slides=len(order))
    out(slides=len(order), body=len(used), prompts=prompts)


def repair(path):
    """에이전트가 쓴 답 JSON 을 확인 · 고침(앞뒤 설명 · 코드 울타리 · 닫힘 실수). 못 고치면 빈 답으로 두고 목록에 남김."""
    if not os.path.exists(path):
        return f"{os.path.basename(path)} 없음"
    raw = open(path, encoding="utf-8-sig").read()
    try:
        json.loads(raw); return None
    except ValueError:
        pass
    text = re.sub(r"```(json)?", "", raw)
    try:
        i = min(k for k in (text.find("{"), text.find("[")) if k >= 0)
        o, end = json.JSONDecoder().raw_decode(text[i:])
        rest = text[i + end:].strip()
        if isinstance(o, dict) and rest.startswith(","):
            try:
                o = json.loads(text[i:i + end - 1] + rest)
            except ValueError:
                pass
        jsave(path, o)
        return None
    except (ValueError, StopIteration):
        jsave(path, {"texts": {}, "checks": [{"slide": 0, "note": f"{os.path.basename(path)} 답을 못 읽어 그 묶음은 원문 그대로 둠"}]})
        return f"{os.path.basename(path)} 못 읽음"


# ---------------- apply · fix ----------------
def over_from_boxes(rows):
    res = []
    for r in rows if isinstance(rows, list) else []:
        if "bw" not in r or str(r.get("t", "")).startswith("확인 ·"):
            continue
        if r["bh"] > r["h"] + 4 or (r.get("wrap") == 0 and r["bw"] > r["w"] + 4):
            res.append({"slide": r["slide"], "id": r["id"], "w": r["w"], "h": r["h"], "bw": r["bw"], "bh": r["bh"], "text": r["t"]})
    return res


def measure(c, post, deck, tag, prev=None, slides=None):
    """넘침 · 겹침 재기. prev(앞 box 파일)와 slides(고친 장 번호)를 주면 그 장만 다시 재고 나머지는 앞 결과를 씀."""
    if ppt_busy():
        raise SystemExit(json.dumps({"error": "PowerPoint 가 열려 있어 멈춤"}, ensure_ascii=False))
    for _ in range(600):  # 원본 자리 재기(뒤에서 도는 것)가 끝나길 기다림
        if os.path.exists(c.p("box_base.json")) and os.path.getsize(c.p("box_base.json")) > 10:
            break
        time.sleep(1)
    # 뒤에서 재다 끊긴 경우(PowerPoint 를 같이 쓰다 실패 등) — 장이 모자라면 원본을 다시 잼
    from pptx import Presentation
    n_slides = len(Presentation(c.p("merged.pptx")).slides)
    base_rows = post.load(c.p("box_base.json")) or []
    base_rows = base_rows if isinstance(base_rows, list) else [base_rows]
    if len({r.get("slide") for r in base_rows}) < n_slides - 2:  # 표지 · 끝장처럼 글 없는 장 두 장까지는 봐줌
        ps("boxes.ps1", "-File", c.p("merged.pptx"), "-Out", c.p("box_base.json"))
        if os.path.exists(c.p("over_base.json")):
            os.remove(c.p("over_base.json"))
    if not os.path.exists(c.p("over_base.json")):
        jsave(c.p("over_base.json"), over_from_boxes(post.load(c.p("box_base.json"))))
    bp = c.p(f"box_{tag}.json")
    if prev and slides and os.path.exists(prev):
        ps("boxes.ps1", "-File", deck, "-Out", bp, "-Slides", ",".join(map(str, sorted(slides))))
        fresh = post.load(bp) or []
        fresh = fresh if isinstance(fresh, list) else [fresh]
        keep = [r for r in post.load(prev) if r["slide"] not in slides]
        jsave(bp, keep + fresh)
    else:
        ps("boxes.ps1", "-File", deck, "-Out", bp)
    op, vp = c.p(f"over_{tag}.json"), c.p(f"ovl_{tag}.json")
    jsave(op, over_from_boxes(post.load(bp)))
    import contextlib, io
    with contextlib.redirect_stdout(io.StringIO()):
        post.overlaps(c.p("box_base.json"), bp, vp)
    new_over = post.overflow_new(post.load(op), post.load(c.p("over_base.json")))
    ovl = [x for x in post.load(vp) if x["area"] >= 60]
    return op, vp, new_over, ovl


def fix_prompt(c, post, deck, op, vp, k):
    import contextlib, io
    with contextlib.redirect_stdout(io.StringIO()):
        post.fixprompt(deck, op, c.p("over_base.json"), c.p(f"fixA{k}.md"))
        post.fixprompt_ovl(deck, vp, c.p(f"fixB{k}.md"))
    a = open(c.p(f"fixA{k}.md"), encoding="utf-8").read()
    b = open(c.p(f"fixB{k}.md"), encoding="utf-8").read()
    rows = list(dict.fromkeys(re.findall(r"^\d+:\d+:\d+ \|.*$", a + "\n" + b, re.M)))
    head = a.split("# 줄일 항목")[0].replace("도형 안에서 글이 넘쳤다", "도형 안에서 글이 넘치거나 옆 도형을 덮었다")
    txt = head + "# 줄일 항목 (id | 최대 글자 | 줄 수 | 지금 글)\n" + "\n".join(rows) + f'\n\n# 출력\n{c.p(f"fix_{k}.json")} 에 JSON 하나만 쓴다: {{"texts": {{"id": "줄인 글", ...}}}}'
    open(c.p(f"fix_{k}.md"), "w", encoding="utf-8").write(txt)
    return len(rows)


def apply(data, work):
    c = Ctx(data, work)
    draft, post, merge = c.mods()
    t = time.time()
    import contextlib, io, glob
    bad = [b for b in (repair(f) for f in sorted(glob.glob(c.p("ans_*.json")))) if b]
    with contextlib.redirect_stdout(io.StringIO()):
        draft.apply(c.p("merged.pptx"), "ans_*.json", c.p("d1.pptx"))
    op, vp, new_over, ovl = measure(c, post, c.p("d1.pptx"), "p1")
    n = 0
    if new_over or ovl:
        with contextlib.redirect_stdout(io.StringIO()):
            post.fixwrap(c.p("d1.pptx"), op, c.p("over_base.json"), c.p("d1.pptx"))
        n = fix_prompt(c, post, c.p("d1.pptx"), op, vp, 1)
    jsave(c.p("state.json"), {"deck": c.p("d1.pptx"), "op": op, "vp": vp})
    log_step(c, "넣기 · 재기", t, over=len(new_over), ovl=len(ovl))
    out(over=len(new_over), overlap=len({(x["slide"], x["id"]) for x in ovl}), fix=(c.p("fix_1.md") if n else None), fix_answer=(c.p("fix_1.json") if n else None), bad_answers=bad)


def fix(data, work, k):
    c = Ctx(data, work)
    draft, post, merge = c.mods()
    k = int(k)
    t = time.time()
    st = jload(c.p("state.json"))
    nxt = c.p(f"d{k + 1}.pptx")
    import contextlib, io
    repair(c.p(f"fix_{k}.json"))
    with contextlib.redirect_stdout(io.StringIO()):
        post.apply_texts(st["deck"], c.p(f"fix_{k}.json"), nxt)
    fixed = {int(key.split(":")[0]) for key in (jload(c.p(f"fix_{k}.json"), {}) or {}).get("texts", {}) if key.split(":")[0].isdigit()}
    op, vp, new_over, ovl = measure(c, post, nxt, f"p{k + 1}", prev=c.p(f"box_p{k}.json"), slides=fixed)
    n = 0
    if (new_over or ovl) and k < 2:
        n = fix_prompt(c, post, nxt, op, vp, k + 1)
    jsave(c.p("state.json"), {"deck": nxt, "op": op, "vp": vp})
    log_step(c, f"줄여 쓰기 {k}", t, over=len(new_over), ovl=len(ovl))
    out(over=len(new_over), overlap=len({(x["slide"], x["id"]) for x in ovl}), fix=(c.p(f"fix_{k + 1}.md") if n else None), fix_answer=(c.p(f"fix_{k + 1}.json") if n else None))


# ---------------- names · finish ----------------
def names(data, work):
    c = Ctx(data, work)
    draft, post, merge = c.mods()
    st = jload(c.p("state.json"))
    hits = post.scan_names(st["deck"])
    if not hits:
        out(left=0, prompt=None); return
    rows = [f"{k} | 최대 {len(tx) + 4}자 | 줄 {tx.count(chr(10)) + 1} | 남은 이름 {', '.join(nm)} | {json.dumps(tx, ensure_ascii=False)}" for k, tx, nm in hits]
    txt = (f"너는 HCG 제안서 Draft 를 다듬는 컨설턴트다. 아래 문단에 다른 고객사(HCG 과거 고객) 이름이 남아 있다. 이번 고객은 {c.cfg.get('short')}.\n"
           "이름을 지우고 뜻은 살려 고쳐 쓴다. 사례 · 벤치마킹 회사 나열이면 「A社 · B社」처럼 익명으로 하거나 「국내 주요 제조사」 같은 일반 표현으로. 길이는 「최대 N자」 안, 줄 수 그대로.\n\n"
           "# 고칠 문단 (id | 최대 글자 | 줄 수 | 남은 이름 | 지금 글)\n" + "\n".join(rows) + f'\n\n# 출력\n{c.p("names.json")} 에 JSON 하나만 쓴다: {{"texts": {{"id": "고친 글", ...}}}}')
    open(c.p("names.md"), "w", encoding="utf-8").write(txt)
    out(left=len(hits), prompt=c.p("names.md"), answer=c.p("names.json"))


def finish(data, work, dest=""):
    c = Ctx(data, work)
    draft, post, merge = c.mods()
    t = time.time()
    st = jload(c.p("state.json"))
    deck, op, vp = st["deck"], st["op"], st["vp"]
    import contextlib, io
    sink = io.StringIO()
    with contextlib.redirect_stdout(sink):
        if os.path.exists(c.p("names.json")) and not repair(c.p("names.json")):
            post.apply_texts(deck, c.p("names.json"), c.p("dn.pptx")); deck = c.p("dn.pptx")
        png, crops = c.p("png"), c.p("crops")
        for d in (png, crops):
            shutil.rmtree(d, ignore_errors=True)
        if ppt_busy():
            raise SystemExit(json.dumps({"error": "PowerPoint 가 열려 있어 멈춤"}, ensure_ascii=False))
        ps("render.ps1", "-File", deck, "-Dir", png)
        post.pics(deck, png, crops)
        if os.path.isdir(crops) and os.listdir(crops):
            ps("ocr.ps1", "-Dir", crops, "-Out", c.p("ocr.json"))
        else:
            jsave(c.p("ocr.json"), {})
        extra = [{"slide": x["slide"], "id": x["id"], "tag": "글 넘침"} for x in post.overflow_new(post.load(op), post.load(c.p("over_base.json")))]
        extra += [{"slide": x["slide"], "id": x["id"], "tag": "글 겹침"} for x in post.load(vp) if x["area"] >= 60]
        jsave(c.p("marks_extra.json"), extra)
        post.stick(deck, c.p("ocr.json"), c.p("stuck.pptx"), c.p("marks_extra.json"))
        name = f"{c.cfg.get('short', '고객')}_{c.cfg.get('title', '')}_Draft.pptx".replace(" ", "")
        final = c.p(name)
        left = post.props(c.p("stuck.pptx"), final)
    checks = post.load(c.p("checks.json")) or []
    for part, nm in (left or {}).items():
        m = re.search(r"slide(\d+)\.xml", part)
        checks.append({"slide": int(m.group(1)) if m else 0, "note": f"다른 고객 이름 남음({', '.join(nm)}) — {part}", "kind": "흔적"})
    for x in post.overflow_new(post.load(op), post.load(c.p("over_base.json"))):
        checks.append({"slide": x["slide"], "note": f"글 넘침(「{x['text'][:14]}…」) — 줄이거나 도형 키우기", "kind": "넘침"})
    for x in {(o["slide"], o["id"]): o for o in post.load(vp) if o["area"] >= 60}.values():
        checks.append({"slide": x["slide"], "note": f"글이 옆 도형에 닿음(「{x['text'][:14]}…」) — 줄이거나 도형 키우기", "kind": "겹침"})
    for ch in checks:
        ch.setdefault("kind", "로고" if ch["note"].startswith(("다른 고객(", "표지 고객 로고")) else "Claude")
    checks = list({(ch["slide"], ch["note"]): ch for ch in checks}.values())
    rank = {"흔적": 0, "넘침": 1, "겹침": 1, "그림": 2, "로고": 2, "값": 3, "차트": 3, "Claude": 4}
    checks.sort(key=lambda ch: (ch["slide"], rank.get(ch["kind"], 5)))
    jsave(c.p("checks_final.json"), checks)
    md_name = f"{c.cfg.get('short', '고객')}_확인목록.md"
    md = [f"# {c.cfg.get('short')} Draft 확인 목록 ({len(checks)}건)", "", "장 위 노란 딱지(이름이 「HCG확인」으로 시작)는 손본 뒤 지우세요.", ""]
    md += [f"- [ ] {ch['slide']}장 · {ch['kind']} — {ch['note']}" for ch in checks]
    open(c.p(md_name), "w", encoding="utf-8").write("\n".join(md))
    files = [final, c.p(md_name)]
    if dest:
        os.makedirs(dest, exist_ok=True)
        files = [shutil.copy2(f, os.path.join(dest, os.path.basename(f))) for f in files]
    log_step(c, "딱지 · 속성 · 목록", t)
    lg = jload(c.p("run_log.json"), {})
    out(file=files[0], checks_file=files[1], checks=len(checks), kinds=dict(collections.Counter(ch["kind"] for ch in checks)),
        names_left=sum(len(v) for v in (left or {}).values()), sec=round(sum(s["sec"] for s in lg.get("steps", [])), 1))


if __name__ == "__main__":
    cmd, a = sys.argv[1], sys.argv[2:]
    fn = {"setup": setup, "start": start, "docs": docs_cmd, "plan": plan, "plan-check": plan_check, "plan-table": plan_table, "build": build, "apply": apply, "fix": fix, "names": names, "finish": finish}.get(cmd)
    if not fn:
        out(error="단계: setup · start · docs · plan · plan-check · plan-table · build · apply · fix · names · finish"); sys.exit(1)
    fn(*a)
