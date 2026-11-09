# DART 원문에서 고객 숫자 읽기(판 3, N1-B): 공시 검색 → 최근 사업보고서 → 요약재무정보 · 직원 현황 · 최대주주 칸 읽기.
# 금융감독원 전자공시 누리집(dart.fss.or.kr)을 사람이 보는 것과 같은 주소로 읽음. 열쇠 · 따로 연결 없음. 건당 4~5쪽만 읽음.
import json, re, sys, time, http.cookiejar, urllib.parse, urllib.request
from lxml import html as LH

BASE = "https://dart.fss.or.kr"
cj = http.cookiejar.CookieJar()
op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
op.addheaders = [("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")]


def get(url, data=None, ref=None):
    req = urllib.request.Request(url, data=data, headers={"Referer": ref or BASE, "X-Requested-With": "XMLHttpRequest"} if data else {})
    return op.open(req, timeout=20).read().decode("utf-8", "ignore")


def rows(h):
    doc = LH.fromstring(h)
    return [[" ".join(c.text_content().split()) for c in tr if c.tag in ("td", "th")] for tr in doc.iter("tr")]


def num(s):
    s = s.replace(",", "").strip()
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else None


def main(name, out):
    t0 = time.time()
    get(BASE + "/dsab001/main.do")
    data = urllib.parse.urlencode({"textCrpNm": name, "currentPage": 1, "maxResults": 15, "maxLinks": 10, "sort": "date", "series": "desc",
                                   "startDate": time.strftime("%Y%m%d", time.localtime(time.time() - 400 * 86400)), "endDate": time.strftime("%Y%m%d"),
                                   "publicType": "A001", "finalReport": "recent"}).encode()
    h = get(BASE + "/dsab001/search.ax", data, BASE + "/dsab001/main.do")
    hits = []  # (접수번호, 보고서 이름) — 회사 칸에 이름이 있고 보고서가 사업보고서인 줄만
    for tr in LH.fromstring(h).iter("tr"):
        cells = [" ".join(c.text_content().split()) for c in tr if c.tag == "td"]
        links = [a.get("href") or "" for a in tr.iter("a")]
        rcps = [m.group(1) for l in links for m in [re.search(r"rcpNo=(\d+)", l)] if m]
        if len(cells) >= 3 and rcps and "사업보고서" in cells[2] and name.replace(" ", "") in cells[1].replace(" ", ""):
            hits.append((rcps[0], cells[2]))
    res = {"listed": bool(hits), "source": "DART 원문", "items": [], "note": ""}
    if not hits:
        res["note"] = "DART 정기공시에서 사업보고서를 못 찾음(비상장이거나 이름이 다름) → 모두 빈 값"
        json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1); print(res["note"]); return
    rcp = hits[0][0]
    res["report"] = f"{' '.join(hits[0][1].split())} · 접수번호 {rcp}"
    main_h = get(f"{BASE}/dsaf001/main.do?rcpNo={rcp}")
    nodes = re.findall(r"node\d+\['text'\]\s*=\s*\"([^\"]*)\";.*?\['dcmNo'\]\s*=\s*\"(\d+)\";.*?\['eleId'\]\s*=\s*\"(\d+)\";.*?\['offset'\]\s*=\s*\"(\d+)\";.*?\['length'\]\s*=\s*\"(\d+)\";", main_h, re.S)
    def sec(word):
        for t, d, e, o, l in nodes:
            if word in t:
                url = f"{BASE}/report/viewer.do?rcpNo={rcp}&dcmNo={d}&eleId={e}&offset={o}&length={l}&dtd=dart4.xsd"
                return get(url), url
        return None, None
    src = f"{BASE}/dsaf001/main.do?rcpNo={rcp}"
    h, url = sec("요약재무정보")
    if h:
        unit = 1000 if "천원" in h else (1000000 if "백만원" in h else 1)
        rr = rows(h)
        years = next(([re.search(r"(\d{4})", c).group(1) for c in r[1:] if re.search(r"\d{4}", c)] for r in rr if r and r[0] == "구 분" and "~" in " ".join(r)), [])
        if not years:  # 머리 칸이 「제42기 · 제41기 …」면 보고서 이름의 연도(2025.12)에서 거꾸로
            m = re.search(r"\((\d{4})\.\d+\)", res["report"])
            ncol = max((len(r) - 1 for r in rr if r and r[0] in ("매출액", "영업이익")), default=0)
            years = [str(int(m.group(1)) - k) for k in range(ncol)] if m else []
        for r in rr:
            if r and r[0] in ("매출액", "영업이익") and years:
                for y, v in zip(years, r[1:]):
                    n = num(v)
                    if n is None or any(i["key"] == r[0] and i["year"] == int(y) for i in res["items"]):
                        continue  # 연결 · 별도 두 표가 있으면 앞(연결) 것만
                    basis = "연결" if "연결" in h[:h.find("매출액")] else "별도"
                    res["items"].append({"key": r[0], "year": int(y), "value": round(n * unit / 1e8, 1), "unit": f"억 원({basis})", "source": src})
    h, url = sec("임원 및 직원 등의 현황")
    if h:
        for r in rows(h):
            if r and r[0] == "합 계" and len(r) >= 6 and num(r[5]) is not None:
                res["items"].append({"key": "직원 수", "year": None, "value": int(num(r[5])), "unit": f"명(기간 정함 없음 {r[1]} · 기간제 {r[3]})", "source": src})
                break
        m = re.search(r"(\d{4})년\s*(\d+)월\s*(\d+)일", " ".join(" ".join(r) for r in rows(h) if r and "기준일" in r[0]))
        if m and res["items"] and res["items"][-1]["key"] == "직원 수":
            res["items"][-1]["year"] = int(m.group(1)); res["items"][-1]["unit"] += f" · 기준일 {m.group(1)}.{m.group(2)}.{m.group(3)}"
    h, url = sec("주주에 관한 사항")
    if h:
        for r in rows(h):
            if len(r) > 5 and r[1:2] and r[1] in ("최대주주", "본인"):
                pct = [c for c in r if re.fullmatch(r"\d+\.\d+", c)]
                res["items"].append({"key": "최대주주", "year": None, "value": f"{r[0]} {pct[-1] + '%' if pct else ''}".strip(), "unit": "기말 지분율", "source": src})
                break
    res["sec"] = round(time.time() - t0, 1)
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(res, ensure_ascii=False)[:1500])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
