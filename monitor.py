# -*- coding: utf-8 -*-
"""
국회 유통 관련 법안 모니터링
- 열린국회정보 OpenAPI(국회사무처)에서 법률안·입법예고를 수집
- config.json 의 키워드로 관련 법안만 추려 Excel/HTML 보고서 생성
- 직전 실행 결과와 비교해 '신규 발의'와 '진행상황 변동'을 표시

사용:
    python monitor.py            # 수집 + 보고서 생성
    python monitor.py --check    # API 연결/응답 항목 점검 (문제 생길 때)
"""
import argparse
import html
import json
import os
import ssl
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

KST = timezone(timedelta(hours=9))


def now_kst():
    return datetime.now(KST)


try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_URL = "https://open.assembly.go.kr/portal/openapi/"
HERE = os.path.dirname(os.path.abspath(__file__))

# 활용 API (오픈API 활용가이드 2023.11 기준)
API_BILL_SEARCH = "TVBPMBILL11"         # API-01-054 법률안 심사 및 처리(의안검색) – 정부·위원장안 포함 전체
API_MEMBER_BILL = "nzmimeepazxkubdpn"   # API-01-006 국회의원 발의법률안 – 대표발의자·위원회 처리결과 보강
API_NOTICE_ING  = "nknalejkafmvgzmpt"   # API-01-052 진행중 입법예고

PAGE_SIZE = 1000

# 응답 항목명이 API마다 조금씩 달라 후보를 순서대로 찾음
FIELD_CANDIDATES = {
    "bill_id":       ["BILL_ID"],
    "bill_no":       ["BILL_NO"],
    "name":          ["BILL_NAME", "BILL_NM"],
    "proposer_kind": ["PROPOSER_KIND", "PROPOSER_KIND_CD"],
    "proposer":      ["PROPOSER"],
    "rep_proposer":  ["RST_PROPOSER"],
    "propose_dt":    ["PROPOSE_DT", "PPSL_DT"],
    "committee":     ["CURR_COMMITTEE", "COMMITTEE", "COMMITTEE_NM"],
    "committee_dt":  ["COMMITTEE_DT"],
    "cmt_result":    ["CMT_PROC_RESULT_CD", "COMMITTEE_PROC_RESULT"],
    "cmt_proc_dt":   ["CMT_PROC_DT", "COMMITTEE_PROC_DT"],
    "law_result":    ["LAW_PROC_RESULT_CD"],
    "proc_result":   ["PROC_RESULT_CD", "PROC_RESULT"],
    "proc_dt":       ["PROC_DT"],
    "link":          ["LINK_URL", "DETAIL_LINK"],
    "notice_end":    ["NOTI_ED_DT", "NOTI_END_DT"],
}

# 변동 감지 대상 항목
TRACKED_FIELDS = [
    ("committee", "소관위"),
    ("cmt_result", "위원회 처리"),
    ("law_result", "법사위 처리"),
    ("proc_result", "본회의 처리"),
    ("proc_dt", "의결일"),
]


# ───────────────────────── API 호출 ─────────────────────────
class ApiError(Exception):
    pass


def _ssl_context(verify):
    if verify:
        return ssl.create_default_context()
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def http_get_json(url, verify_ssl=True, retries=3):
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (assembly-monitor)"})
            with urllib.request.urlopen(req, timeout=60, context=_ssl_context(verify_ssl)) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:  # 네트워크 일시 오류 재시도
            last = e
            time.sleep(1.5 * (i + 1))
    raise ApiError(f"접속 실패: {last}")


def call_api(cfg, api, params, p_index=1, p_size=PAGE_SIZE):
    """반환: (전체건수, rows, 결과코드, 메시지)"""
    q = {"KEY": cfg["api_key"], "Type": "json", "pIndex": p_index, "pSize": p_size}
    q.update({k: v for k, v in params.items() if v not in (None, "")})
    url = BASE_URL + api + "?" + urllib.parse.urlencode(q)
    data = http_get_json(url, cfg.get("verify_ssl", True))

    # 데이터 없음/오류는 최상위 RESULT 로 옴
    if "RESULT" in data:
        res = data["RESULT"]
        return 0, [], res.get("CODE", ""), res.get("MESSAGE", "")

    body = data.get(api)
    if not body:
        raise ApiError(f"예상하지 못한 응답 형식: {str(data)[:300]}")
    total, code, msg, rows = 0, "", "", []
    for part in body:
        if "head" in part:
            for h in part["head"]:
                if "list_total_count" in h:
                    total = int(h["list_total_count"])
                if "RESULT" in h:
                    code = h["RESULT"].get("CODE", "")
                    msg = h["RESULT"].get("MESSAGE", "")
        if "row" in part:
            rows = part["row"]
    return total, rows, code, msg


def fetch_all(cfg, api, params, label):
    rows_all, p = [], 1
    while True:
        total, rows, code, msg = call_api(cfg, api, params, p)
        if code and code not in ("INFO-000", "INFO-200"):
            raise ApiError(f"{label}: {code} {msg}")
        rows_all.extend(rows)
        print(f"  · {label}: {len(rows_all):,} / {total:,}건", end="\r")
        if not rows or len(rows_all) >= total:
            break
        p += 1
        time.sleep(0.2)
    print(f"  · {label}: {len(rows_all):,}건 수집 완료        ")
    return rows_all


# ───────────────────────── 정규화·필터 ─────────────────────────
def pick(row, key):
    for f in FIELD_CANDIDATES[key]:
        v = row.get(f)
        if v not in (None, ""):
            return str(v).strip()
    return ""


def normalize(row):
    return {k: pick(row, k) for k in FIELD_CANDIDATES}


def _squash(s):
    return "".join(str(s).split())


def match_groups(name, committee, cfg):
    """의안명이 키워드에 걸리면 [(그룹, 키워드)] 반환"""
    n = _squash(name)
    if any(_squash(x) in n for x in cfg.get("exclude_keywords", []) if x):
        return []
    hits = []
    for group, kws in cfg["keyword_groups"].items():
        for kw in kws:
            if kw and _squash(kw) in n:
                hits.append((group, kw))
                break
    if not hits and committee and committee in cfg.get("watch_committees", []):
        hits.append(("관심 위원회", committee))
    return hits


def status_of(b):
    if b.get("proc_result"):
        return b["proc_result"]
    if b.get("law_result"):
        return f"법사위 {b['law_result']}"
    if b.get("cmt_result"):
        return f"소관위 {b['cmt_result']}"
    if b.get("committee"):
        return "소관위 심사"
    return "접수"


def stage_of(b):
    if b.get("proc_result"):
        return 4
    if b.get("law_result"):
        return 3
    if b.get("cmt_result"):
        return 2
    return 1


# ───────────────────────── 수집 ─────────────────────────
def collect(cfg):
    age = str(cfg.get("assembly_age", 22))
    bills = {}

    print("① 의안검색(전체 법률안) 수집")
    for r in fetch_all(cfg, API_BILL_SEARCH, {"AGE": age}, "의안검색"):
        b = normalize(r)
        if b["bill_id"]:
            bills[b["bill_id"]] = b

    print("② 의원 발의법률안 수집(대표발의자·위원회 처리결과 보강)")
    try:
        for r in fetch_all(cfg, API_MEMBER_BILL, {"AGE": age}, "의원발의"):
            m = normalize(r)
            if not m["bill_id"]:
                continue
            base = bills.setdefault(m["bill_id"], m)
            for k, v in m.items():
                if v and not base.get(k):
                    base[k] = v
    except ApiError as e:
        print(f"  ! 의원발의 API 보강 생략: {e}")

    tracked = []
    for b in bills.values():
        hits = match_groups(b["name"], b["committee"], cfg)
        if hits:
            b["group"] = ", ".join(sorted({g for g, _ in hits}))
            b["keyword"] = ", ".join(sorted({k for _, k in hits}))
            if not b["proposer_kind"] and b["rep_proposer"]:
                b["proposer_kind"] = "의원"
            b["status"] = status_of(b)
            tracked.append(b)

    print("③ 진행중 입법예고 수집")
    notices = []
    try:
        try:
            raw = fetch_all(cfg, API_NOTICE_ING, {}, "입법예고")
        except ApiError:
            raw = fetch_all(cfg, API_NOTICE_ING, {"AGE": age}, "입법예고")
        for r in raw:
            n = normalize(r)
            hits = match_groups(n["name"], n["committee"], cfg)
            if hits:
                n["group"] = ", ".join(sorted({g for g, _ in hits}))
                notices.append(n)
    except ApiError as e:
        print(f"  ! 입법예고 수집 실패: {e}")

    tracked.sort(key=lambda x: (x.get("propose_dt", ""), x.get("bill_no", "")), reverse=True)
    notices.sort(key=lambda x: x.get("notice_end", ""))
    return tracked, notices


# ───────────────────────── 직전 결과와 비교 ─────────────────────────
def diff(tracked, snapshot_path, recent_days):
    prev = {}
    first_run = not os.path.exists(snapshot_path)
    if not first_run:
        with open(snapshot_path, encoding="utf-8") as f:
            prev = json.load(f).get("bills", {})

    cutoff = (now_kst() - timedelta(days=recent_days)).strftime("%Y-%m-%d")
    events = []
    for b in tracked:
        old = prev.get(b["bill_id"])
        b["flag"] = ""
        if first_run:
            if b.get("propose_dt", "") >= cutoff:
                b["flag"] = "최근발의"
                events.append({"구분": f"최근 {recent_days}일 발의", "bill": b, "내용": "발의/제출"})
            continue
        if old is None:
            b["flag"] = "신규"
            events.append({"구분": "신규 발의", "bill": b, "내용": "발의/제출"})
            continue
        changes = []
        for key, label in TRACKED_FIELDS:
            if (old.get(key) or "") != (b.get(key) or ""):
                changes.append(f"{label}: {old.get(key) or '-'} → {b.get(key) or '-'}")
        if changes:
            b["flag"] = "변동"
            events.append({"구분": "진행상황 변동", "bill": b, "내용": " / ".join(changes)})
    return events, first_run


def save_snapshot(tracked, path):
    keep = ["bill_no", "name", "committee", "cmt_result", "law_result", "proc_result", "proc_dt"]
    data = {
        "saved_at": now_kst().strftime("%Y-%m-%d %H:%M"),
        "bills": {b["bill_id"]: {k: b.get(k, "") for k in keep} for b in tracked},
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


# ───────────────────────── Excel 보고서 ─────────────────────────
NAVY, BLUE, LIGHT = "1F3864", "2E75B6", "DEEAF6"

BILL_COLS = [
    ("구분", "flag", 8), ("관심분야", "group", 18), ("의안번호", "bill_no", 10),
    ("의안명", "name", 46), ("제안자구분", "proposer_kind", 10), ("대표발의/제안자", "_proposer", 18),
    ("제안일", "propose_dt", 11), ("소관위원회", "committee", 26), ("진행상태", "status", 16),
    ("의결일", "proc_dt", 11), ("의안 링크", "link", 12),
]
NOTICE_COLS = [
    ("관심분야", "group", 18), ("의안번호", "bill_no", 10), ("의안명", "name", 46),
    ("제안자", "proposer", 22), ("소관위원회", "committee", 26), ("예고 종료일", "notice_end", 12),
    ("의견 등록 링크", "link", 14),
]


def _proposer(b):
    return b.get("rep_proposer") or b.get("proposer") or ""


def _val(b, key):
    return _proposer(b) if key == "_proposer" else b.get(key, "")


def write_excel(path, tracked, notices, events, meta):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor=NAVY)
    head_font = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
    body_font = Font(name="맑은 고딕", size=10)
    link_font = Font(name="맑은 고딕", size=10, color=BLUE, underline="single")
    flag_fill = {"신규": PatternFill("solid", fgColor="FCE4D6"),
                 "변동": PatternFill("solid", fgColor="FFF2CC"),
                 "최근발의": PatternFill("solid", fgColor="E2EFDA")}

    wb = Workbook()

    # 요약
    ws = wb.active
    ws.title = "요약"
    ws["A1"] = "국회 유통 관련 법안 모니터링 보고"
    ws["A1"].font = Font(name="맑은 고딕", size=16, bold=True, color=NAVY)
    rows = [
        ("작성일시", meta["run_at"]),
        ("대상", f"제{meta['age']}대 국회 법률안 (의안명 키워드 기준)"),
        ("비교 기준", meta["compare"]),
        ("추적 법안", f"{len(tracked):,}건"),
        ("신규 발의", f"{meta['n_new']:,}건"),
        ("진행상황 변동", f"{meta['n_changed']:,}건"),
        ("진행중 입법예고", f"{len(notices):,}건"),
        ("자료 출처", "열린국회정보 OpenAPI (국회사무처)"),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        a, b = ws.cell(i, 1, k), ws.cell(i, 2, v)
        a.font, a.fill = Font(name="맑은 고딕", size=10, bold=True, color=NAVY), PatternFill("solid", fgColor=LIGHT)
        b.font = body_font
        a.border = b.border = border
    r0 = len(rows) + 5
    ws.cell(r0, 1, "관심분야별 현황").font = Font(name="맑은 고딕", size=12, bold=True, color=NAVY)
    hdr = ["관심분야", "추적 법안", "계류", "처리완료", "입법예고 진행"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r0 + 1, j, h)
        c.font, c.fill, c.border = head_font, head_fill, border
        c.alignment = Alignment(horizontal="center")
    for i, g in enumerate(meta["groups"], start=r0 + 2):
        gb = [b for b in tracked if g in b.get("group", "")]
        vals = [g, len(gb), sum(1 for b in gb if not b.get("proc_result")),
                sum(1 for b in gb if b.get("proc_result")),
                sum(1 for n in notices if g in n.get("group", ""))]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            c.font, c.border = body_font, border
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 44
    for col in "CDE":
        ws.column_dimensions[col].width = 14

    def table(ws, cols, data, flag_key=None):
        for j, (h, _, w) in enumerate(cols, 1):
            c = ws.cell(1, j, h)
            c.font, c.fill, c.border = head_font, head_fill, border
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(j)].width = w
        for i, item in enumerate(data, start=2):
            b = item["bill"] if "bill" in item else item
            for j, (h, key, _) in enumerate(cols, 1):
                v = item.get(key) if key in ("구분", "내용") else _val(b, key)
                c = ws.cell(i, j)
                if key == "link" and v:
                    c.value, c.hyperlink, c.font = "열기", v, link_font
                else:
                    c.value, c.font = v, body_font
                c.border = border
                c.alignment = Alignment(vertical="top", wrap_text=key in ("name", "내용", "group"))
            if flag_key:
                f = b.get(flag_key)
                if f in flag_fill:
                    for j in range(1, len(cols) + 1):
                        ws.cell(i, j).fill = flag_fill[f]
        ws.freeze_panes = "A2"
        if data:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(data) + 1}"

    ev_cols = [("구분", "구분", 14), ("내용", "내용", 40)] + [c for c in BILL_COLS if c[1] != "flag"]
    table(wb.create_sheet("신규·변동"), ev_cols, events, "flag")
    table(wb.create_sheet("추적법안 전체"), BILL_COLS, tracked, "flag")
    table(wb.create_sheet("진행중 입법예고"), NOTICE_COLS, notices)
    wb.save(path)


def write_csv_fallback(path, tracked):
    import csv
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([h for h, _, _ in BILL_COLS])
        for b in tracked:
            w.writerow([_val(b, k) for _, k, _ in BILL_COLS])


# ───────────────────────── HTML 보고서 ─────────────────────────
def write_html(path, tracked, notices, events, meta, nav=""):
    e = html.escape

    def link(u, t="열기"):
        return f'<a href="{e(u)}" target="_blank">{t}</a>' if u else ""

    def bill_rows(items, with_event=False):
        out = []
        for it in items:
            b = it["bill"] if with_event else it
            cls = {"신규": "new", "변동": "chg", "최근발의": "recent"}.get(b.get("flag", ""), "")
            ev = f"<td class=\"nw\">{e(it['구분'])}</td><td class=\"ev\">{e(it['내용'])}</td>" if with_event else f"<td>{e(b.get('flag',''))}</td>"
            out.append(
                f'<tr class="{cls}">{ev}<td>{e(b.get("group",""))}</td><td>{e(b.get("bill_no",""))}</td>'
                f'<td class="nm">{e(b.get("name",""))}</td><td>{e(_proposer(b))}</td>'
                f'<td>{e(b.get("propose_dt",""))}</td><td>{e(b.get("committee",""))}</td>'
                f'<td><b>{e(b.get("status",""))}</b></td><td>{link(b.get("link",""))}</td></tr>')
        return "\n".join(out)

    head_ev = "<th>구분</th><th>내용</th>"
    head_b = "<th>관심분야</th><th>의안번호</th><th>의안명</th><th>대표발의/제안자</th><th>제안일</th><th>소관위</th><th>진행상태</th><th>링크</th>"
    notice_rows = "\n".join(
        f'<tr><td>{e(n.get("group",""))}</td><td>{e(n.get("bill_no",""))}</td><td class="nm">{e(n.get("name",""))}</td>'
        f'<td>{e(n.get("proposer",""))}</td><td>{e(n.get("committee",""))}</td><td><b>{e(n.get("notice_end",""))}</b></td>'
        f'<td>{link(n.get("link",""), "의견등록")}</td></tr>' for n in notices)
    empty = '<p class="muted">해당 없음</p>'

    doc = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>국회 법안 모니터링 {e(meta['run_at'])}</title>
<style>
body{{font-family:'맑은 고딕','Malgun Gothic',sans-serif;margin:24px;color:#222;background:#fff}}
h1{{color:#{NAVY};border-bottom:3px solid #{NAVY};padding-bottom:6px;font-size:22px}}
h2{{color:#{NAVY};font-size:17px;margin-top:28px;border-left:5px solid #{BLUE};padding-left:8px}}
table{{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px}}
th{{background:#{NAVY};color:#fff;padding:6px;border:1px solid #9aa;white-space:nowrap}}
td{{border:1px solid #ccc;padding:5px;vertical-align:top}} td.nm{{min-width:260px}}
td.ev{{min-width:200px}} td.nw{{white-space:nowrap}} td:nth-child(n){{word-break:keep-all}}
.meta td{{border:1px solid #ccc}} .meta td:first-child{{background:#{LIGHT};font-weight:bold;color:#{NAVY};width:140px}}
tr.new td{{background:#FCE4D6}} tr.chg td{{background:#FFF2CC}} tr.recent td{{background:#E2EFDA}}
.muted{{color:#888;font-size:13px}} a{{color:#{BLUE}}}
.wrap{{overflow-x:auto;-webkit-overflow-scrolling:touch}}
.nav{{display:flex;gap:14px;flex-wrap:wrap;font-size:14px;margin-bottom:6px}}
.nav a{{background:#{LIGHT};padding:6px 10px;border-radius:4px;text-decoration:none;font-weight:bold}}
@media (max-width:700px){{body{{margin:12px}} h1{{font-size:19px}} table{{font-size:12px}} td.nm{{min-width:200px}}}}
</style></head><body>
{nav}
<h1>국회 유통 관련 법안 모니터링 보고</h1>
<table class="meta">
<tr><td>작성일시</td><td>{e(meta['run_at'])}</td></tr>
<tr><td>대상</td><td>제{meta['age']}대 국회 법률안 (의안명 키워드 기준)</td></tr>
<tr><td>비교 기준</td><td>{e(meta['compare'])}</td></tr>
<tr><td>요약</td><td>추적 {len(tracked):,}건 · 신규 {meta['n_new']}건 · 변동 {meta['n_changed']}건 · 진행중 입법예고 {len(notices)}건</td></tr>
</table>
<h2>1. 신규 발의 및 진행상황 변동</h2>
{f'<div class="wrap"><table><tr>{head_ev}{head_b}</tr>{bill_rows(events, True)}</table></div>' if events else empty}
<h2>2. 진행중 입법예고 (예고 종료일 순)</h2>
{f'<div class="wrap"><table><tr><th>관심분야</th><th>의안번호</th><th>의안명</th><th>제안자</th><th>소관위</th><th>예고 종료일</th><th>링크</th></tr>{notice_rows}</table></div>' if notices else empty}
<h2>3. 추적 법안 전체 (제안일 최신순)</h2>
{f'<div class="wrap"><table><tr><th>구분</th>{head_b}</tr>{bill_rows(tracked)}</table></div>' if tracked else empty}
<p class="muted">자료: 열린국회정보 OpenAPI(국회사무처). 키워드는 의안명 기준으로 판별하므로 본문만 관련된 법안은 누락될 수 있음.</p>
</body></html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


# ───────────────────────── 점검 모드 ─────────────────────────
def check(cfg):
    age = str(cfg.get("assembly_age", 22))
    for api, label, params in [
        (API_BILL_SEARCH, "의안검색", {"AGE": age}),
        (API_MEMBER_BILL, "의원발의", {"AGE": age}),
        (API_NOTICE_ING, "진행중 입법예고", {}),
    ]:
        try:
            total, rows, code, msg = call_api(cfg, api, params, 1, 1)
            print(f"[{label}] {api} → {code} {msg} / 전체 {total:,}건")
            if rows:
                print("   응답 항목:", ", ".join(rows[0].keys()))
        except Exception as ex:
            print(f"[{label}] {api} → 실패: {ex}")


# ───────────────────────── main ─────────────────────────
def load_config():
    with open(os.path.join(HERE, "config.json"), encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["api_key"] = os.environ.get("ASSEMBLY_API_KEY") or cfg.get("api_key", "")
    if not cfg["api_key"]:
        sys.exit("인증키가 없습니다. config.json 의 api_key 또는 ASSEMBLY_API_KEY 환경변수를 설정하세요.")
    return cfg


def run(cfg, out_dir=None, state_dir=None, site_dir=None, keep=90):
    out_dir = out_dir or os.path.join(HERE, "보고서")
    state_dir = state_dir or os.path.join(HERE, "data")
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(state_dir, exist_ok=True)
    snap = os.path.join(state_dir, "snapshot.json")

    prev_time = ""
    if os.path.exists(snap):
        with open(snap, encoding="utf-8") as f:
            prev_time = json.load(f).get("saved_at", "")

    tracked, notices = collect(cfg)
    events, first = diff(tracked, snap, cfg.get("recent_days", 14))
    events.sort(key=lambda x: ({"신규 발의": 0, "진행상황 변동": 1}.get(x["구분"], 2), x["bill"].get("propose_dt", "")))

    now = now_kst()
    meta = {
        "run_at": now.strftime("%Y-%m-%d %H:%M"),
        "age": cfg.get("assembly_age", 22),
        "compare": (f"최초 실행 – 최근 {cfg.get('recent_days', 14)}일 발의 법안 표시" if first
                    else f"직전 실행({prev_time}) 대비"),
        "n_new": sum(1 for x in events if x["구분"] != "진행상황 변동"),
        "n_changed": sum(1 for x in events if x["구분"] == "진행상황 변동"),
        "groups": list(cfg["keyword_groups"].keys()) + (["관심 위원회"] if cfg.get("watch_committees") else []),
    }
    if site_dir:
        out_dir = os.path.join(site_dir, "reports")
        os.makedirs(out_dir, exist_ok=True)
        name = f"report_{now:%Y%m%d_%H%M}"
    else:
        name = f"국회법안모니터링_{now:%Y%m%d_%H%M}"
    stem = os.path.join(out_dir, name)
    try:
        write_excel(stem + ".xlsx", tracked, notices, events, meta)
        xl = stem + ".xlsx"
    except ImportError:
        write_csv_fallback(stem + ".csv", tracked)
        xl = stem + ".csv (openpyxl 미설치로 CSV 저장)"
    write_html(stem + ".html", tracked, notices, events, meta,
               nav=('<div class="nav"><a href="../index.html">최신 보고서</a>'
                    '<a href="../archive.html">지난 보고서</a>'
                    f'<a href="{name}.xlsx">엑셀 받기</a></div>') if site_dir else "")
    if site_dir:
        write_site(site_dir, name, tracked, notices, events, meta, keep)
    save_snapshot(tracked, snap)

    print("\n완료")
    print(f"  추적 {len(tracked):,}건 | 신규 {meta['n_new']}건 | 변동 {meta['n_changed']}건 | 입법예고 {len(notices)}건")
    print(f"  Excel: {xl}")
    print(f"  HTML : {stem}.html")
    return stem


def write_site(site_dir, name, tracked, notices, events, meta, keep):
    """웹 게시용: index.html(최신), archive.html(지난 보고서), 오래된 보고서 정리"""
    rep = os.path.join(site_dir, "reports")
    stems = sorted({f.rsplit(".", 1)[0] for f in os.listdir(rep) if f.startswith("report_")}, reverse=True)
    for old in stems[keep:]:
        for ext in (".html", ".xlsx", ".csv"):
            p = os.path.join(rep, old + ext)
            if os.path.exists(p):
                os.remove(p)
    stems = stems[:keep]

    write_html(os.path.join(site_dir, "index.html"), tracked, notices, events, meta,
               nav=('<div class="nav"><a href="archive.html">지난 보고서</a>'
                    f'<a href="reports/{name}.xlsx">엑셀 받기</a></div>'))

    items = []
    for s in stems:
        d = s.replace("report_", "")
        label = f"{d[0:4]}-{d[4:6]}-{d[6:8]} {d[9:11]}:{d[11:13]}"
        items.append(f'<li><a href="reports/{s}.html">{label}</a> · <a href="reports/{s}.xlsx">엑셀</a></li>')
    with open(os.path.join(site_dir, "archive.html"), "w", encoding="utf-8") as f:
        f.write(f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>지난 보고서</title>
<style>body{{font-family:'맑은 고딕','Malgun Gothic',sans-serif;margin:20px;background:#fff;color:#222}}
h1{{color:#{NAVY};font-size:20px;border-bottom:3px solid #{NAVY};padding-bottom:6px}}
li{{margin:10px 0;font-size:15px}} a{{color:#{BLUE}}}</style></head><body>
<p><a href="index.html">← 최신 보고서</a></p><h1>지난 보고서</h1><ul>{''.join(items)}</ul></body></html>""")
    open(os.path.join(site_dir, ".nojekyll"), "w").close()


def main():
    ap = argparse.ArgumentParser(description="국회 유통 관련 법안 모니터링")
    ap.add_argument("--check", action="store_true", help="API 연결 및 응답 항목 점검")
    ap.add_argument("--site", metavar="DIR", help="웹 게시용 폴더(예: docs)에 index.html 등 생성")
    args = ap.parse_args()
    cfg = load_config()
    if args.check:
        check(cfg)
        return
    try:
        if args.site:
            run(cfg, state_dir=os.path.join(HERE, "data"), site_dir=os.path.join(HERE, args.site))
        else:
            run(cfg)
    except ApiError as ex:
        print(f"\n오류: {ex}\n'python monitor.py --check' 로 연결 상태를 확인하세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()
