from flask import Flask, render_template_string, jsonify
import requests
import csv
import io
import re

app = Flask(__name__)

SERIES_URL = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=0"
SERIES_NAVER_URL = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=1734309078"
CVR_URL    = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=478441460"
ALRIM_URL    = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=650933119"
CUSTOMER_URL = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=1173300152"
MARKETING_URL = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=30238607"
NPS_URL       = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=537236483"
STRATEGY_URL  = "https://docs.google.com/spreadsheets/d/1VbJPL4g1ItSc6JKKAhtyj0JTLC62MD7iiqct533AGfY/export?format=csv&gid=706038710"


def fetch_series():
    resp = requests.get(SERIES_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    headers = rows[0][1:]
    weeks = {}
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        label = row[0].strip()
        match = re.match(r"(.+?)[(\s](실적|비중)[)\s]*$", label)
        if not match:
            continue
        week_name = match.group(1).strip()
        kind = match.group(2)
        values = row[1:]
        if week_name not in weeks:
            weeks[week_name] = {}
        weeks[week_name][kind] = values
    result = []
    for week_name, data in weeks.items():
        entry = {"week": week_name, "brands": []}
        perf_list  = data.get("실적", [])
        ratio_list = data.get("비중", [])
        for i, brand in enumerate(headers):
            perf_raw  = perf_list[i]  if i < len(perf_list)  else ""
            ratio_raw = ratio_list[i] if i < len(ratio_list) else ""
            perf_num  = int(perf_raw.replace(",", "").replace('"', "")) if perf_raw.replace(",", "").replace('"', "").lstrip("-").isdigit() else None
            ratio_str = ratio_raw.strip().replace('%', '') if ratio_raw.strip() else None
            entry["brands"].append({"brand": brand, "perf": perf_num, "ratio": ratio_str})
        result.append(entry)
    return result


def parse_num(s):
    if not s or not s.strip():
        return None
    clean = s.strip().replace(',', '').replace('%', '').replace('"', '')
    try:
        return float(clean)
    except ValueError:
        return None


def fetch_series_naver():
    resp = requests.get(SERIES_NAVER_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    headers = rows[0][1:]
    weeks = {}
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        label = row[0].strip()
        match = re.match(r"(.+?)[(\s](실적|비중)[)\s]*$", label)
        if not match:
            continue
        week_name = match.group(1).strip()
        kind = match.group(2)
        values = row[1:]
        if week_name not in weeks:
            weeks[week_name] = {}
        weeks[week_name][kind] = values
    result = []
    for week_name, data in weeks.items():
        entry = {"week": week_name, "brands": []}
        perf_list  = data.get("실적", [])
        ratio_list = data.get("비중", [])
        for i, brand in enumerate(headers):
            perf_raw  = perf_list[i]  if i < len(perf_list)  else ""
            ratio_raw = ratio_list[i] if i < len(ratio_list) else ""
            perf_num  = int(perf_raw.replace(",", "").replace('"', "")) if perf_raw.replace(",", "").replace('"', "").lstrip("-").isdigit() else None
            ratio_str = ratio_raw.strip().replace('%', '') if ratio_raw.strip() else None
            entry["brands"].append({"brand": brand, "perf": perf_num, "ratio": ratio_str})
        result.append(entry)
    return result


def fetch_cvr():
    resp = requests.get(CVR_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    result = []
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        week = row[0].strip()
        cvr  = parse_num(row[11]) if len(row) > 11 else None
        roas = parse_num(row[12]) if len(row) > 12 else None
        aov  = parse_num(row[14]) if len(row) > 14 else None
        result.append({"week": week, "cvr": cvr, "roas": roas, "aov": aov})
    return result


def fetch_alrim():
    resp = requests.get(ALRIM_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    # A=주차(0), B=전송대상(1), C=읽음수(2), D=확인율(3), E=클릭수(4), F=클릭율(5), G=주문금액(6), H=전송당매출(7)
    result = []
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        week       = row[0].strip()
        confirm    = parse_num(row[3]) if len(row) > 3 else None
        click_rate = parse_num(row[5]) if len(row) > 5 else None
        order_amt  = parse_num(row[6]) if len(row) > 6 else None
        rev_per    = parse_num(row[7]) if len(row) > 7 else None
        result.append({"week": week, "confirm": confirm, "click_rate": click_rate,
                        "order_amt": order_amt, "rev_per": rev_per})
    return result


def fetch_customer():
    resp = requests.get(CUSTOMER_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    # A=주차(0), B=구매자수(1), C=1년내재구매자수(2), D=1년내최초구매자수(3)
    result = []
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        week    = row[0].strip()
        total   = parse_num(row[1]) if len(row) > 1 else None
        repurch = parse_num(row[2]) if len(row) > 2 else None
        new_    = parse_num(row[3]) if len(row) > 3 else None
        new_rate    = round(new_    / total * 100, 2) if total and new_    is not None else None
        repurch_rate = round(repurch / total * 100, 2) if total and repurch is not None else None
        result.append({
            "week": week,
            "total": total, "new": new_, "repurch": repurch,
            "new_rate": new_rate, "repurch_rate": repurch_rate
        })
    return result


def fetch_marketing():
    resp = requests.get(MARKETING_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    # A=월(0), C=네이버수주(2), D=네이버MER(3), G=온라인수주(6), H=온라인MER(7)
    result = []
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        month       = row[0].strip()
        naver_ord   = parse_num(row[2]) if len(row) > 2 else None
        naver_mer   = parse_num(row[3]) if len(row) > 3 else None
        online_ord  = parse_num(row[6]) if len(row) > 6 else None
        online_mer  = parse_num(row[7]) if len(row) > 7 else None
        result.append({
            "month": month,
            "naver_ord": naver_ord, "naver_mer": naver_mer,
            "online_ord": online_ord, "online_mer": online_mer,
        })
    return result


def fetch_strategy():
    resp = requests.get(STRATEGY_URL, timeout=10)
    resp.encoding = "utf-8"
    rows = list(csv.reader(io.StringIO(resp.text)))
    if len(rows) < 2:
        return {"weeks": [], "data": {}}

    # Row 0: 카테고리, 시리즈, week1, '', '', week2, '', '', ...
    # Row 1: '', '', 오프라인, 온라인, 사업개발, 오프라인, 온라인, 사업개발, ...
    week_cols = []  # list of (week_name, col_index)
    for i, v in enumerate(rows[0]):
        if i >= 2 and v.strip():
            week_cols.append((v.strip(), i))

    # Build {category: {series: {week: {offline, online, biz}}}}
    data = {}
    cat_order = []
    series_order = {}
    for row in rows[2:]:
        if not row or not row[0].strip():
            continue
        cat = row[0].strip()
        ser = row[1].strip()
        if not ser:
            continue
        if cat not in data:
            data[cat] = {}
            cat_order.append(cat)
            series_order[cat] = []
        if ser not in data[cat]:
            data[cat][ser] = {}
            series_order[cat].append(ser)
        for week_name, col in week_cols:
            biz     = parse_num(row[col])     if len(row) > col     else None
            offline = parse_num(row[col + 1]) if len(row) > col + 1 else None
            online  = parse_num(row[col + 2]) if len(row) > col + 2 else None
            if week_name not in data[cat][ser]:
                data[cat][ser][week_name] = {"offline": 0, "online": 0, "biz": 0}
            if offline is not None:
                data[cat][ser][week_name]["offline"] += offline
            if online is not None:
                data[cat][ser][week_name]["online"]  += online
            if biz is not None:
                data[cat][ser][week_name]["biz"]     += biz

    weeks = [w for w, _ in week_cols]
    return {"weeks": weeks, "categories": cat_order, "series_order": series_order, "data": data}


def fetch_nps():
    resp = requests.get(NPS_URL, timeout=10)
    resp.encoding = "utf-8"
    reader = csv.reader(io.StringIO(resp.text))
    rows = list(reader)
    # A=월(0), B=추천자(1), C=중립자(2), D=비방자(3), E=NPS점수(4), F=전체NPS점수(5)
    result = []
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        month    = row[0].strip()
        promo    = parse_num(row[1]) if len(row) > 1 else None
        neutral  = parse_num(row[2]) if len(row) > 2 else None
        detract  = parse_num(row[3]) if len(row) > 3 else None
        nps      = parse_num(row[4]) if len(row) > 4 else None
        nps_total = parse_num(row[5]) if len(row) > 5 else None
        total    = (promo or 0) + (neutral or 0) + (detract or 0)
        result.append({
            "month": month,
            "promo": promo, "neutral": neutral, "detract": detract,
            "promo_pct":   round(promo   / total * 100, 1) if total else None,
            "neutral_pct": round(neutral / total * 100, 1) if total else None,
            "detract_pct": round(detract / total * 100, 1) if total else None,
            "nps": nps,
            "nps_total": nps_total,
        })
    return result


HTML = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>KGI 대시보드</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Malgun Gothic', sans-serif; background: #f0f2f5; color: #333; }
  header {
    background: #1a237e; color: white;
    padding: 16px 24px;
    display: flex; align-items: center; justify-content: space-between;
  }
  header h1 { font-size: 20px; }
  #refresh-btn {
    background: #fff; color: #1a237e;
    border: none; padding: 8px 18px; border-radius: 6px;
    font-size: 14px; cursor: pointer; font-weight: bold;
  }
  #refresh-btn:hover { background: #e8eaf6; }
  #status { font-size: 12px; color: #aaa; padding: 8px 24px; }
  .container { padding: 16px 24px; }

  .radio-tabs {
    display: flex; gap: 0; margin-bottom: 20px;
    background: white; border-radius: 10px; overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1); width: fit-content;
  }
  .radio-tabs label {
    padding: 12px 32px; cursor: pointer; font-size: 12pt; font-weight: bold;
    color: #888; border-right: 1px solid #eee; transition: all 0.15s;
  }
  .radio-tabs label:last-child { border-right: none; }
  .radio-tabs input[type=radio] { display: none; }
  .radio-tabs input[type=radio]:checked + label { background: #1a237e; color: white; }

  .section { display: none; }
  .section.active { display: block; }

  .week-selector {
    background: white; border-radius: 10px; padding: 14px 18px;
    margin-bottom: 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.1);
  }
  .week-selector h3 { font-size: 13px; color: #666; margin-bottom: 10px; }
  .week-checks { display: flex; gap: 10px; flex-wrap: wrap; }
  .week-check-label {
    display: flex; align-items: center; gap: 6px;
    padding: 6px 14px; border-radius: 20px; cursor: pointer;
    border: 2px solid #ccc; font-size: 13px; user-select: none;
    transition: all 0.15s;
  }
  .week-check-label.checked { color: white; border-color: transparent; }
  .week-check-label input { display: none; }

  .totals-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
  .total-card {
    background: white; border-radius: 8px; padding: 10px 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    border-left: 4px solid #ccc; min-width: 160px;
  }
  .total-card .wname { font-size: 13px; color: #888; }
  .total-card .wamt  { font-size: 15px; font-weight: bold; margin-top: 2px; }

  .sub-tabs { display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }
  .sub-tab {
    padding: 8px 22px; border-radius: 6px; cursor: pointer; font-size: 14px;
    background: white; border: 2px solid #ccc; font-weight: bold; color: #888;
    transition: all 0.15s;
  }
  .sub-tab.active { background: #1a237e; color: white; border-color: #1a237e; }

  .chart-wrap {
    background: white; border-radius: 10px; padding: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1);
  }
</style>
</head>
<body>
<header>
  <h1>KGI 대시보드</h1>
  <button id="refresh-btn" onclick="loadAll()">새로고침</button>
</header>
<div id="status">불러오는 중...</div>
<div class="container">

  <div class="radio-tabs">
    <input type="radio" name="dash" id="r-strategy" checked onchange="switchDash('strategy')">
    <label for="r-strategy">사업부 실적</label>
    <input type="radio" name="dash" id="r-cat-perf" onchange="switchDash('cat-perf')">
    <label for="r-cat-perf">카테고리실적</label>
    <input type="radio" name="dash" id="r-series-naver" onchange="switchDash('series-naver')">
    <label for="r-series-naver">시리즈별실적(네이버)</label>
    <input type="radio" name="dash" id="r-cvr" onchange="switchDash('cvr')">
    <label for="r-cvr">전환율</label>
    <input type="radio" name="dash" id="r-alrim" onchange="switchDash('alrim')">
    <label for="r-alrim">알림받기전환</label>
    <input type="radio" name="dash" id="r-customer" onchange="switchDash('customer')">
    <label for="r-customer">신규고객율/재구매고객율</label>
    <input type="radio" name="dash" id="r-marketing" onchange="switchDash('marketing')">
    <label for="r-marketing">마케팅(월)</label>
    <input type="radio" name="dash" id="r-nps" onchange="switchDash('nps')">
    <label for="r-nps">NPS(월)</label>
  </div>


  <!-- 시리즈별실적(네이버) -->
  <div id="sec-series-naver" class="section">
    <div class="week-selector">
      <h3>주차 선택</h3>
      <div class="week-checks" id="series-naver-week-checks"></div>
    </div>
    <div class="totals-row" id="series-naver-totals"></div>
    <div class="chart-wrap"><canvas id="series-naver-chart"></canvas></div>
  </div>

  <!-- 전환율 -->
  <div id="sec-cvr" class="section">
    <div class="sub-tabs">
      <div class="sub-tab active" id="cvr-tab-cvr"  onclick="setCvrSub('cvr')">CVR (전환율)</div>
      <div class="sub-tab"        id="cvr-tab-roas" onclick="setCvrSub('roas')">ROAS (광고대비매출액)</div>
      <div class="sub-tab"        id="cvr-tab-aov"  onclick="setCvrSub('aov')">AOV (객단가)</div>
    </div>
    <div class="chart-wrap"><canvas id="cvr-chart"></canvas></div>
  </div>

  <!-- 신규고객/재구매고객율 -->
  <div id="sec-customer" class="section">
    <div class="sub-tabs">
      <div class="sub-tab active" id="customer-tab-new"     onclick="setCustomerSub('new')">신규구매</div>
      <div class="sub-tab"        id="customer-tab-repurch" onclick="setCustomerSub('repurch')">재구매</div>
    </div>
    <div class="chart-wrap"><canvas id="customer-chart"></canvas></div>
  </div>

  <!-- 마케팅(월) -->
  <div id="sec-marketing" class="section">
    <div class="sub-tabs">
      <div class="sub-tab active" id="mkt-tab-naver"  onclick="setMktSub('naver')">네이버</div>
      <div class="sub-tab"        id="mkt-tab-online" onclick="setMktSub('online')">온라인</div>
    </div>
    <div class="chart-wrap"><canvas id="marketing-chart"></canvas></div>
  </div>

  <!-- NPS -->
  <div id="sec-nps" class="section">
    <div class="chart-wrap"><canvas id="nps-chart"></canvas></div>
  </div>

  <!-- 사업부 실적 -->
  <div id="sec-strategy" class="section active">
    <div class="sub-tabs" id="strategy-cat-tabs"></div>
    <div class="week-selector">
      <h3>시리즈 선택</h3>
      <div class="week-checks" id="strategy-series-checks"></div>
    </div>
    <div class="week-selector">
      <h3>채널 선택</h3>
      <div class="week-checks" id="strategy-channel-checks"></div>
    </div>
    <div class="totals-row" id="strategy-totals"></div>
    <div class="chart-wrap"><canvas id="strategy-chart"></canvas></div>
  </div>

  <!-- 카테고리실적 -->
  <div id="sec-cat-perf" class="section">
    <div class="chart-wrap"><canvas id="cat-perf-chart"></canvas></div>
  </div>

  <!-- 알림받기전환 -->
  <div id="sec-alrim" class="section">
    <div class="sub-tabs">
      <div class="sub-tab active" id="alrim-tab-rate"  onclick="setAlrimSub('rate')">확인율 / 클릭율</div>
      <div class="sub-tab"        id="alrim-tab-sales" onclick="setAlrimSub('sales')">주문금액 / 전송당매출</div>
    </div>
    <div class="chart-wrap"><canvas id="alrim-chart"></canvas></div>
  </div>

</div>

<script>
let seriesData   = [];
let seriesNaverData = [];
let cvrData      = [];
let alrimData      = [];
let customerData   = [];
let marketingData  = [];
let selectedWeeks = new Set();
let selectedNaverWeeks = new Set();
let seriesChart   = null;
let seriesNaverChart = null;
let cvrChart      = null;
let alrimChart    = null;
let customerChart  = null;
let marketingChart = null;
let npsChart       = null;
let npsData        = [];
let strategyData   = {};
let strategyChart  = null;
let catPerfChart   = null;
let catPerfCat     = '';
let strategyCat    = '';
let selectedStrategySeries = new Set();
let selectedStrategyChannels = new Set(['offline', 'online', 'biz']);
let cvrSub      = 'cvr';
let alrimSub    = 'rate';
let customerSub = 'new';
let mktSub      = 'naver';

const COLORS = [
  '#1a237e','#e53935','#2e7d32','#f57c00','#6a1b9a',
  '#00838f','#558b2f','#ad1457','#4527a0','#00695c'
];

async function loadAll() {
  document.getElementById('status').textContent = '불러오는 중...';
  try {
    const [r2, r3, r4, r5, r6, r7, r8] = await Promise.all([
      fetch('/api/cvr'), fetch('/api/alrim'),
      fetch('/api/customer'), fetch('/api/marketing'), fetch('/api/nps'),
      fetch('/api/series_naver'), fetch('/api/strategy')
    ]);
    cvrData       = await r2.json();
    alrimData     = await r3.json();
    customerData  = await r4.json();
    marketingData = await r5.json();
    npsData       = await r6.json();
    seriesNaverData = await r7.json();
    strategyData  = await r8.json();
    selectedNaverWeeks = new Set(seriesNaverData.slice(0, 2).map(w => w.week));
    renderSeriesNaverWeekSelector();
    renderSeriesNaver();
    renderCvr();
    renderAlrim();
    renderCustomer();
    renderMarketing();
    renderNps();
    renderStrategyInit();
    renderCatPerfInit();
    const now = new Date().toLocaleString('ko-KR');
    document.getElementById('status').textContent = '마지막 갱신: ' + now;
  } catch(e) {
    document.getElementById('status').textContent = '오류: ' + e.message;
  }
}

function switchDash(name) {
  ['series-naver','cvr','alrim','customer','marketing','nps','strategy','cat-perf'].forEach(n => {
    document.getElementById('sec-' + n).className = 'section' + (n === name ? ' active' : '');
  });
}

/* ===== 시리즈별실적 ===== */
function renderSeriesWeekSelector() {
  const container = document.getElementById('series-week-checks');
  container.innerHTML = seriesData.map((w, i) => {
    const color   = COLORS[i % COLORS.length];
    const checked = selectedWeeks.has(w.week);
    return `<label class="week-check-label ${checked?'checked':''}" style="${checked?'background:'+color+';border-color:'+color:''}">
      <input type="checkbox" ${checked?'checked':''} onchange="toggleSeriesWeek('${w.week}',this,'${color}',this.parentElement)">
      ${w.week}
    </label>`;
  }).join('');
}

function toggleSeriesWeek(weekName, cb, color, label) {
  if (cb.checked) {
    selectedWeeks.add(weekName);
    label.classList.add('checked');
    label.style.background  = color;
    label.style.borderColor = color;
  } else {
    selectedWeeks.delete(weekName);
    label.classList.remove('checked');
    label.style.background  = '';
    label.style.borderColor = '';
  }
  renderSeries();
}

function fmtNum(n) {
  if (n === null || n === undefined) return '-';
  return n.toLocaleString('ko-KR');
}

function renderSeries() {
  if (!seriesData.length) return;
  const active = seriesData.filter(w => selectedWeeks.has(w.week));

  document.getElementById('series-totals').innerHTML = active.map(w => {
    const color = COLORS[seriesData.indexOf(w) % COLORS.length];
    const total = w.brands.find(b => b.brand === '합계');
    return `<div class="total-card" style="border-left-color:${color}">
      <div class="wname">${w.week}</div>
      <div class="wamt" style="color:${color}">${total ? fmtNum(total.perf) : '-'}개</div>
    </div>`;
  }).join('');

  const brandSet = new Set();
  for (const w of seriesData)
    for (const b of w.brands)
      if (b.brand !== '합계' && b.ratio && parseFloat(b.ratio) > 0) brandSet.add(b.brand);
  const brands = Array.from(brandSet);

  const allAvg = brands.map(brand => {
    const vals = seriesData.map(w => {
      const b = w.brands.find(x => x.brand === brand);
      return b && b.ratio ? parseFloat(b.ratio) : 0;
    });
    return parseFloat((vals.reduce((a,b)=>a+b,0) / seriesData.length).toFixed(1));
  });

  const datasets = active.map(w => {
    const color = COLORS[seriesData.indexOf(w) % COLORS.length];
    const data  = brands.map(brand => {
      const b = w.brands.find(x => x.brand === brand);
      return b && b.ratio ? parseFloat(b.ratio) : null;
    });
    return {
      label: w.week, data,
      borderColor: color, backgroundColor: color,
      pointBackgroundColor: color, pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 8, pointHoverRadius: 11, showLine: false,
      datalabels: {
        align: 'top', offset: 4, color,
        font: { size: 13, weight: 'bold' },
        formatter: v => v !== null ? parseFloat(v).toFixed(1) + '%' : '',
      }
    };
  });

  datasets.push({
    label: '전체 평균', data: allAvg,
    backgroundColor: '#9e9e9e', pointBackgroundColor: '#9e9e9e',
    pointBorderColor: '#fff', pointBorderWidth: 2,
    pointRadius: 6, pointHoverRadius: 9, pointStyle: 'rectRot', showLine: false,
    datalabels: { display: false }
  });

  if (seriesChart) seriesChart.destroy();
  seriesChart = new Chart(document.getElementById('series-chart').getContext('2d'), {
    type: 'line',
    data: { labels: brands, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const label = context.dataset.label;
              const val   = context.parsed.y;
              if (label === '전체 평균') return ` ${label}: ${parseFloat(val).toFixed(1)}%`;
              const w = active[context.datasetIndex];
              if (!w) return '';
              const b = w.brands.find(x => x.brand === brands[context.dataIndex]);
              const perf = b && b.perf !== null ? fmtNum(b.perf) + '개' : '-';
              return ` ${label}  비중: ${parseFloat(val).toFixed(1)}%  /  실적: ${perf}`;
            }
          }
        }
      },
      scales: {
        y: { title: { display: true, text: '비중 (%)' }, ticks: { callback: v => v + '%' }, min: 0 },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== 시리즈별실적(네이버) ===== */
function renderSeriesNaverWeekSelector() {
  const container = document.getElementById('series-naver-week-checks');
  container.innerHTML = seriesNaverData.map((w, i) => {
    const color   = COLORS[i % COLORS.length];
    const checked = selectedNaverWeeks.has(w.week);
    return `<label class="week-check-label ${checked?'checked':''}" style="${checked?'background:'+color+';border-color:'+color:''}">
      <input type="checkbox" ${checked?'checked':''} onchange="toggleSeriesNaverWeek('${w.week}',this,'${color}',this.parentElement)">
      ${w.week}
    </label>`;
  }).join('');
}

function toggleSeriesNaverWeek(weekName, cb, color, label) {
  if (cb.checked) {
    selectedNaverWeeks.add(weekName);
    label.classList.add('checked');
    label.style.background  = color;
    label.style.borderColor = color;
  } else {
    selectedNaverWeeks.delete(weekName);
    label.classList.remove('checked');
    label.style.background  = '';
    label.style.borderColor = '';
  }
  renderSeriesNaver();
}

function renderSeriesNaver() {
  if (!seriesNaverData.length) return;
  const active = seriesNaverData.filter(w => selectedNaverWeeks.has(w.week));

  document.getElementById('series-naver-totals').innerHTML = active.map(w => {
    const color = COLORS[seriesNaverData.indexOf(w) % COLORS.length];
    const total = w.brands.find(b => b.brand === '합계');
    return `<div class="total-card" style="border-left-color:${color}">
      <div class="wname">${w.week}</div>
      <div class="wamt" style="color:${color}">${total ? fmtNum(total.perf) : '-'}개</div>
    </div>`;
  }).join('');

  const brandSet = new Set();
  for (const w of seriesNaverData)
    for (const b of w.brands)
      if (b.brand !== '합계' && b.ratio && parseFloat(b.ratio) > 0) brandSet.add(b.brand);
  const brands = Array.from(brandSet);

  const allAvg = brands.map(brand => {
    const vals = seriesNaverData.map(w => {
      const b = w.brands.find(x => x.brand === brand);
      return b && b.ratio ? parseFloat(b.ratio) : 0;
    });
    return parseFloat((vals.reduce((a,b)=>a+b,0) / seriesNaverData.length).toFixed(1));
  });

  const datasets = active.map(w => {
    const color = COLORS[seriesNaverData.indexOf(w) % COLORS.length];
    const data  = brands.map(brand => {
      const b = w.brands.find(x => x.brand === brand);
      return b && b.ratio ? parseFloat(b.ratio) : null;
    });
    return {
      label: w.week, data,
      borderColor: color, backgroundColor: color,
      pointBackgroundColor: color, pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 8, pointHoverRadius: 11, showLine: false,
      datalabels: {
        align: 'top', offset: 4, color,
        font: { size: 13, weight: 'bold' },
        formatter: v => v !== null ? parseFloat(v).toFixed(1) + '%' : '',
      }
    };
  });

  datasets.push({
    label: '전체 평균', data: allAvg,
    backgroundColor: '#9e9e9e', pointBackgroundColor: '#9e9e9e',
    pointBorderColor: '#fff', pointBorderWidth: 2,
    pointRadius: 6, pointHoverRadius: 9, pointStyle: 'rectRot', showLine: false,
    datalabels: { display: false }
  });

  if (seriesNaverChart) seriesNaverChart.destroy();
  seriesNaverChart = new Chart(document.getElementById('series-naver-chart').getContext('2d'), {
    type: 'line',
    data: { labels: brands, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const label = context.dataset.label;
              const val   = context.parsed.y;
              if (label === '전체 평균') return ` ${label}: ${parseFloat(val).toFixed(1)}%`;
              const w = active[context.datasetIndex];
              if (!w) return '';
              const b = w.brands.find(x => x.brand === brands[context.dataIndex]);
              const perf = b && b.perf !== null ? fmtNum(b.perf) + '개' : '-';
              return ` ${label}  비중: ${parseFloat(val).toFixed(1)}%  /  실적: ${perf}`;
            }
          }
        }
      },
      scales: {
        y: { title: { display: true, text: '비중 (%)' }, ticks: { callback: v => v + '%' }, min: 0 },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== 전환율 ===== */
function setCvrSub(sub) {
  cvrSub = sub;
  ['cvr','roas','aov'].forEach(s => {
    document.getElementById('cvr-tab-' + s).className = 'sub-tab' + (s === sub ? ' active' : '');
  });
  renderCvr();
}

function renderCvr() {
  if (!cvrData.length) return;
  const meta = {
    cvr:  { key: 'cvr',  label: 'CVR (전환율)',         unit: '%',  color: '#1a237e', fmt: v => v.toFixed(2) + '%' },
    roas: { key: 'roas', label: 'ROAS (광고대비매출액)', unit: '%',  color: '#2e7d32', fmt: v => v.toFixed(0) + '%' },
    aov:  { key: 'aov',  label: 'AOV (객단가)',           unit: '원', color: '#e53935', fmt: v => v.toLocaleString('ko-KR') + '원' },
  }[cvrSub];

  const sorted = [...cvrData].reverse();
  const labels = sorted.map(r => r.week);
  const values = sorted.map(r => r[meta.key]);
  const valid  = values.filter(v => v !== null);
  const avg    = valid.length ? parseFloat((valid.reduce((a,b)=>a+b,0) / valid.length).toFixed(2)) : 0;

  const datasets = [
    {
      label: meta.label, data: values,
      backgroundColor: meta.color, pointBackgroundColor: meta.color,
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 8, pointHoverRadius: 11, showLine: false,
      datalabels: {
        align: 'top', offset: 4, color: meta.color,
        font: { size: 13, weight: 'bold' },
        formatter: v => v !== null ? meta.fmt(v) : '',
      }
    },
    {
      label: '전체 평균', data: labels.map(() => avg),
      backgroundColor: '#9e9e9e', pointBackgroundColor: '#9e9e9e',
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 6, pointHoverRadius: 9, pointStyle: 'rectRot', showLine: false,
      datalabels: { display: false }
    }
  ];

  if (cvrChart) cvrChart.destroy();
  cvrChart = new Chart(document.getElementById('cvr-chart').getContext('2d'), {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const label = context.dataset.label;
              const val   = context.parsed.y;
              if (label === '전체 평균') return ` 전체 평균: ${meta.fmt(val)}`;
              return ` ${label}: ${meta.fmt(val)}`;
            }
          }
        }
      },
      scales: {
        y: { title: { display: true, text: meta.label }, ticks: { callback: v => v + meta.unit } },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== 알림받기전환 ===== */
function setAlrimSub(sub) {
  alrimSub = sub;
  ['rate','sales'].forEach(s => {
    document.getElementById('alrim-tab-' + s).className = 'sub-tab' + (s === sub ? ' active' : '');
  });
  renderAlrim();
}

function calcAvg(arr) {
  const valid = arr.filter(v => v !== null && v !== undefined);
  return valid.length ? valid.reduce((a,b) => a+b, 0) / valid.length : 0;
}

function renderAlrim() {
  if (!alrimData.length) return;
  const sorted = [...alrimData];
  const labels = sorted.map(r => r.week);

  if (alrimSub === 'rate') {
    // 확인율(%) + 클릭율(%) — 같은 Y축
    const confirmVals    = sorted.map(r => r.confirm);
    const clickVals      = sorted.map(r => r.click_rate);
    const confirmAvg     = parseFloat(calcAvg(confirmVals).toFixed(2));
    const clickAvg       = parseFloat(calcAvg(clickVals).toFixed(2));

    const datasets = [
      {
        label: '확인율', data: confirmVals,
        backgroundColor: '#1a237e', pointBackgroundColor: '#1a237e',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 8, pointHoverRadius: 11, showLine: false,
        yAxisID: 'y',
        datalabels: {
          align: 'top', offset: 4, color: '#1a237e',
          font: { size: 13, weight: 'bold' },
          formatter: v => v !== null ? v.toFixed(2) + '%' : '',
        }
      },
      {
        label: '클릭율', data: clickVals,
        backgroundColor: '#e53935', pointBackgroundColor: '#e53935',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 8, pointHoverRadius: 11, showLine: false,
        yAxisID: 'y',
        datalabels: {
          align: 'bottom', offset: 4, color: '#e53935',
          font: { size: 13, weight: 'bold' },
          formatter: v => v !== null ? v.toFixed(2) + '%' : '',
        }
      },
      {
        label: '확인율 평균', data: labels.map(() => confirmAvg),
        backgroundColor: '#7986cb', pointBackgroundColor: '#7986cb',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
        yAxisID: 'y', datalabels: { display: false }
      },
      {
        label: '클릭율 평균', data: labels.map(() => clickAvg),
        backgroundColor: '#ef9a9a', pointBackgroundColor: '#ef9a9a',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
        yAxisID: 'y', datalabels: { display: false }
      }
    ];

    if (alrimChart) alrimChart.destroy();
    alrimChart = new Chart(document.getElementById('alrim-chart').getContext('2d'), {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top' },
          datalabels: { display: true },
          tooltip: {
            callbacks: {
              label(context) {
                const label = context.dataset.label;
                const val   = context.parsed.y;
                return ` ${label}: ${val !== null ? val.toFixed(2) + '%' : '-'}`;
              }
            }
          }
        },
        scales: {
          y: { title: { display: true, text: '비율 (%)' }, ticks: { callback: v => v + '%' }, min: 0 },
          x: { ticks: { maxRotation: 45 } }
        }
      },
      plugins: [ChartDataLabels]
    });

  } else {
    // 주문금액(원) + 전송당매출(원) — 단위 차이로 이중 Y축
    const orderVals  = sorted.map(r => r.order_amt);
    const revVals    = sorted.map(r => r.rev_per);
    const orderAvg   = Math.round(calcAvg(orderVals));
    const revAvg     = parseFloat(calcAvg(revVals).toFixed(0));

    const datasets = [
      {
        label: '주문금액', data: orderVals,
        backgroundColor: '#2e7d32', pointBackgroundColor: '#2e7d32',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 8, pointHoverRadius: 11, showLine: false,
        yAxisID: 'y1',
        datalabels: {
          align: 'top', offset: 4, color: '#2e7d32',
          font: { size: 13, weight: 'bold' },
          formatter: v => v !== null ? v.toLocaleString('ko-KR') + '원' : '',
        }
      },
      {
        label: '전송당매출', data: revVals,
        backgroundColor: '#f57c00', pointBackgroundColor: '#f57c00',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 8, pointHoverRadius: 11, showLine: false,
        yAxisID: 'y2',
        datalabels: {
          align: 'bottom', offset: 4, color: '#f57c00',
          font: { size: 13, weight: 'bold' },
          formatter: v => v !== null ? v.toLocaleString('ko-KR') + '원' : '',
        }
      },
      {
        label: '주문금액 평균', data: labels.map(() => orderAvg),
        backgroundColor: '#a5d6a7', pointBackgroundColor: '#a5d6a7',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
        yAxisID: 'y1', datalabels: { display: false }
      },
      {
        label: '전송당매출 평균', data: labels.map(() => revAvg),
        backgroundColor: '#ffcc80', pointBackgroundColor: '#ffcc80',
        pointBorderColor: '#fff', pointBorderWidth: 2,
        pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
        yAxisID: 'y2', datalabels: { display: false }
      }
    ];

    if (alrimChart) alrimChart.destroy();
    alrimChart = new Chart(document.getElementById('alrim-chart').getContext('2d'), {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top' },
          datalabels: { display: true },
          tooltip: {
            callbacks: {
              label(context) {
                const label = context.dataset.label;
                const val   = context.parsed.y;
                return ` ${label}: ${val !== null ? val.toLocaleString('ko-KR') + '원' : '-'}`;
              }
            }
          }
        },
        scales: {
          y1: {
            type: 'linear', position: 'left',
            title: { display: true, text: '주문금액 (원)' },
            ticks: { callback: v => v.toLocaleString('ko-KR') }
          },
          y2: {
            type: 'linear', position: 'right',
            title: { display: true, text: '전송당매출 (원)' },
            ticks: { callback: v => v.toLocaleString('ko-KR') },
            grid: { drawOnChartArea: false }
          },
          x: { ticks: { maxRotation: 45 } }
        }
      },
      plugins: [ChartDataLabels]
    });
  }
}

/* ===== 신규고객/재구매고객율 ===== */
function setCustomerSub(sub) {
  customerSub = sub;
  ['new','repurch'].forEach(s => {
    document.getElementById('customer-tab-' + s).className = 'sub-tab' + (s === sub ? ' active' : '');
  });
  renderCustomer();
}

function renderCustomer() {
  if (!customerData.length) return;
  const isNew   = customerSub === 'new';
  const rateKey = isNew ? 'new_rate' : 'repurch_rate';
  const cntKey  = isNew ? 'new'      : 'repurch';
  const label   = isNew ? '신규구매율 (최초구매/구매자수)' : '재구매율 (재구매/구매자수)';
  const color   = isNew ? '#1a237e' : '#e53935';
  const avgColor = isNew ? '#7986cb' : '#ef9a9a';

  const labels = customerData.map(r => r.week);
  const values = customerData.map(r => r[rateKey]);
  const valid  = values.filter(v => v !== null);
  const avg    = valid.length ? parseFloat((valid.reduce((a,b) => a+b, 0) / valid.length).toFixed(2)) : 0;

  const datasets = [
    {
      label, data: values,
      backgroundColor: color, pointBackgroundColor: color,
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 8, pointHoverRadius: 11, showLine: false,
      datalabels: {
        align: 'top', offset: 4, color,
        font: { size: 13, weight: 'bold' },
        formatter: v => v !== null ? v.toFixed(2) + '%' : '',
      }
    },
    {
      label: '전체 평균', data: labels.map(() => avg),
      backgroundColor: avgColor, pointBackgroundColor: avgColor,
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
      datalabels: { display: false }
    }
  ];

  if (customerChart) customerChart.destroy();
  customerChart = new Chart(document.getElementById('customer-chart').getContext('2d'), {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const lbl = context.dataset.label;
              const val = context.parsed.y;
              if (lbl === '전체 평균') return ` 전체 평균: ${val.toFixed(2)}%`;
              const r = customerData[context.dataIndex];
              return ` ${lbl}: ${val.toFixed(2)}%  (${fmtNum(r[cntKey])}명 / 구매자 ${fmtNum(r.total)}명)`;
            }
          }
        }
      },
      scales: {
        y: { title: { display: true, text: '비율 (%)' }, ticks: { callback: v => v + '%' }, min: 0 },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== 마케팅(월) ===== */
function setMktSub(sub) {
  mktSub = sub;
  ['naver','online'].forEach(s => {
    document.getElementById('mkt-tab-' + s).className = 'sub-tab' + (s === sub ? ' active' : '');
  });
  renderMarketing();
}

function renderMarketing() {
  if (!marketingData.length) return;
  const isNaver  = mktSub === 'naver';
  const ordKey   = isNaver ? 'naver_ord'  : 'online_ord';
  const merKey   = isNaver ? 'naver_mer'  : 'online_mer';
  const ordLabel = isNaver ? '네이버 수주' : '온라인 수주';
  const merLabel = isNaver ? '네이버 MER'  : '온라인 MER';
  const ordColor = isNaver ? '#1a237e' : '#2e7d32';
  const merColor = isNaver ? '#e53935' : '#f57c00';

  const labels   = marketingData.map(r => r.month);
  const ordVals  = marketingData.map(r => r[ordKey]);
  const merVals  = marketingData.map(r => r[merKey]);

  const ordValid = ordVals.filter(v => v !== null);
  const merValid = merVals.filter(v => v !== null);
  const ordAvg   = ordValid.length ? Math.round(ordValid.reduce((a,b)=>a+b,0) / ordValid.length) : 0;
  const merAvg   = merValid.length ? parseFloat((merValid.reduce((a,b)=>a+b,0) / merValid.length).toFixed(1)) : 0;
  const EOK      = 100000000;
  const TWO_EOK  = 200000000;
  const ordMin   = ordValid.length ? Math.floor((Math.min(...ordValid) - TWO_EOK) / EOK) * EOK : 0;
  const ordMax   = ordValid.length ? Math.ceil((Math.max(...ordValid)  + TWO_EOK) / EOK) * EOK : 0;
  const merMin   = merValid.length ? parseFloat((Math.min(...merValid) - 2).toFixed(1)) : 0;
  const merMax   = merValid.length ? parseFloat((Math.max(...merValid) + 2).toFixed(1)) : 0;

  const datasets = [
    {
      label: ordLabel, data: ordVals,
      backgroundColor: ordColor, pointBackgroundColor: ordColor,
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 8, pointHoverRadius: 11, showLine: false,
      yAxisID: 'y1',
      datalabels: {
        align: 'top', offset: 4, color: ordColor,
        font: { size: 13, weight: 'bold' },
        formatter: v => v !== null ? (v / 100000000).toFixed(1) + '억' : '',
      }
    },
    {
      label: merLabel, data: merVals,
      backgroundColor: merColor, pointBackgroundColor: merColor,
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 8, pointHoverRadius: 11, showLine: false,
      yAxisID: 'y2',
      datalabels: {
        align: 'bottom', offset: 4, color: merColor,
        font: { size: 13, weight: 'bold' },
        formatter: v => v !== null ? v.toFixed(1) : '',
      }
    },
    {
      label: ordLabel + ' 평균', data: labels.map(() => ordAvg),
      backgroundColor: ordColor + '88', pointBackgroundColor: ordColor + '88',
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
      yAxisID: 'y1', datalabels: { display: false }
    },
    {
      label: merLabel + ' 평균', data: labels.map(() => merAvg),
      backgroundColor: merColor + '88', pointBackgroundColor: merColor + '88',
      pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 5, pointHoverRadius: 7, pointStyle: 'rectRot', showLine: false,
      yAxisID: 'y2', datalabels: { display: false }
    }
  ];

  if (marketingChart) marketingChart.destroy();
  marketingChart = new Chart(document.getElementById('marketing-chart').getContext('2d'), {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const lbl = context.dataset.label;
              const val = context.parsed.y;
              if (lbl.includes('평균')) return ` ${lbl}: ${lbl.includes(merLabel) ? val.toFixed(1) : fmtNum(Math.round(val)) + '원'}`;
              if (lbl === ordLabel) return ` ${lbl}: ${fmtNum(Math.round(val))}원`;
              return ` ${lbl}: ${val !== null ? val.toFixed(1) : '-'}`;
            }
          }
        }
      },
      scales: {
        y1: {
          type: 'linear', position: 'left',
          min: ordMin, max: ordMax,
          title: { display: true, text: '수주 (원)' },
          ticks: { callback: v => (v / 100000000).toFixed(1) + '억' }
        },
        y2: {
          type: 'linear', position: 'right',
          min: merMin, max: merMax,
          title: { display: true, text: 'MER' },
          ticks: { callback: v => v },
          grid: { drawOnChartArea: false }
        },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== NPS ===== */
function renderNps() {
  if (!npsData.length) return;
  const labels = npsData.map(r => r.month);

  const datasets = [
    {
      type: 'bar',
      label: '추천자',
      data: npsData.map(r => r.promo_pct),
      backgroundColor: '#a8d5a2',
      stack: 'nps',
      datalabels: {
        color: '#fff', font: { size: 13, weight: 'bold' },
        formatter: v => v !== null && v >= 5 ? v.toFixed(1) + '%' : '',
      }
    },
    {
      type: 'bar',
      label: '중립자',
      data: npsData.map(r => r.neutral_pct),
      backgroundColor: '#c9c9c9',
      stack: 'nps',
      datalabels: {
        color: '#fff', font: { size: 13, weight: 'bold' },
        formatter: v => v !== null && v >= 5 ? v.toFixed(1) + '%' : '',
      }
    },
    {
      type: 'bar',
      label: '비방자',
      data: npsData.map(r => r.detract_pct),
      backgroundColor: '#f4a9a8',
      stack: 'nps',
      datalabels: {
        color: '#fff', font: { size: 13, weight: 'bold' },
        formatter: v => v !== null && v >= 5 ? v.toFixed(1) + '%' : '',
      }
    },
    {
      type: 'line',
      label: '온라인 NPS 점수',
      data: npsData.map(r => r.nps),
      borderColor: '#1a237e',
      backgroundColor: '#1a237e',
      pointBackgroundColor: '#1a237e',
      pointBorderColor: '#fff', pointBorderWidth: 3,
      pointRadius: 10, pointHoverRadius: 13,
      showLine: false,
      order: 0,
      yAxisID: 'y2',
      datalabels: {
        align: 'top', offset: 8, color: '#1a237e',
        font: { size: 15, weight: 'bold' },
        formatter: v => v !== null ? '온라인 NPS 점수 : ' + v.toFixed(0) : '',
      }
    },
    {
      type: 'line',
      label: '전체 NPS 점수',
      data: npsData.map(r => r.nps_total),
      borderColor: '#e53935',
      backgroundColor: '#e53935',
      pointBackgroundColor: '#e53935',
      pointBorderColor: '#fff', pointBorderWidth: 3,
      pointRadius: 10, pointHoverRadius: 13,
      showLine: false,
      order: 0,
      yAxisID: 'y2',
      datalabels: {
        align: 'bottom', offset: 8, color: '#e53935',
        font: { size: 15, weight: 'bold' },
        formatter: v => v !== null ? '전체 NPS 점수 : ' + v.toFixed(0) : '',
      }
    }
  ];

  // 최근 3개월 이동평균 추가
  const last3 = npsData.slice(-3);
  const avg3 = key => {
    const vals = last3.map(r => r[key]).filter(v => v !== null);
    return vals.length ? parseFloat((vals.reduce((a,b)=>a+b,0) / vals.length).toFixed(1)) : null;
  };
  const avgLabel = '3개월 이동평균';
  datasets[0].data.push(avg3('promo_pct'));
  datasets[1].data.push(avg3('neutral_pct'));
  datasets[2].data.push(avg3('detract_pct'));
  datasets[3].data.push(avg3('nps'));
  datasets[4].data.push(avg3('nps_total'));
  labels.push(avgLabel);

  if (npsChart) npsChart.destroy();
  npsChart = new Chart(document.getElementById('nps-chart').getContext('2d'), {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const lbl = context.dataset.label;
              const val = context.parsed.y;
              const r   = npsData[context.dataIndex];
              if (lbl === '온라인 NPS 점수') return ` 온라인 NPS 점수: ${val}`;
              if (lbl === '전체 NPS 점수') return ` 전체 NPS 점수: ${val}`;
              const countMap = { '추천자': r.promo, '중립자': r.neutral, '비방자': r.detract };
              return ` ${lbl}: ${val !== null ? val.toFixed(1) + '%' : '-'}  (${fmtNum(countMap[lbl])}명)`;
            }
          }
        }
      },
      scales: {
        y: {
          stacked: true,
          min: 0, max: 100,
          position: 'left',
          ticks: { callback: v => v + '%' },
          title: { display: true, text: '비율 (%)' }
        },
        y2: {
          min: 0, max: 100,
          position: 'right',
          ticks: { callback: v => v },
          title: { display: true, text: '온라인 NPS 점수' },
          grid: { drawOnChartArea: false }
        },
        x: { stacked: true, ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== 사업부 실적 ===== */
// 시리즈별 고유 색상, 채널별 선 스타일로 구분
const STRATEGY_COLORS = [
  '#1a237e','#e53935','#2e7d32','#f57c00','#6a1b9a',
  '#00838f','#558b2f','#ad1457','#4527a0','#00695c',
  '#37474f','#c62828'
];
const CHANNEL_DASH   = { offline: [],        online: [6,3],    biz: [2,2] };
const CHANNEL_POINT  = { offline: 'circle',  online: 'rect',   biz: 'triangle' };
const CHANNEL_LABEL  = { offline: '오프라인', online: '온라인', biz: '사업개발' };
const CHANNEL_ALIGN  = { offline: 'top',     online: 'right',  biz: 'bottom' };

function renderStrategyInit() {
  if (!strategyData || !strategyData.categories || !strategyData.categories.length) return;
  strategyCat = strategyData.categories[0];
  selectedStrategySeries = new Set(strategyData.series_order[strategyCat] || []);
  selectedStrategyChannels = new Set(['offline', 'online', 'biz']);
  renderStrategyCatTabs();
  renderStrategySeriesChecks();
  renderStrategyChannelChecks();
  renderStrategy();
}

function renderStrategyChannelChecks() {
  const container = document.getElementById('strategy-channel-checks');
  const channels = [
    { key: 'offline', label: '오프라인', color: '#1a237e' },
    { key: 'online',  label: '온라인',   color: '#2e7d32' },
    { key: 'biz',     label: '사업개발', color: '#e53935' },
  ];
  container.innerHTML = channels.map(({ key, label, color }) => {
    const checked = selectedStrategyChannels.has(key);
    return `<label class="week-check-label ${checked ? 'checked' : ''}" style="${checked ? 'background:' + color + ';border-color:' + color : ''}">
      <input type="checkbox" ${checked ? 'checked' : ''} onchange="toggleStrategyChannel('${key}', this, '${color}', this.parentElement)">
      ${label}
    </label>`;
  }).join('');
}

function toggleStrategyChannel(key, cb, color, label) {
  if (cb.checked) {
    selectedStrategyChannels.add(key);
    label.classList.add('checked');
    label.style.background  = color;
    label.style.borderColor = color;
  } else {
    selectedStrategyChannels.delete(key);
    label.classList.remove('checked');
    label.style.background  = '';
    label.style.borderColor = '';
  }
  renderStrategy();
}

function renderStrategyCatTabs() {
  const container = document.getElementById('strategy-cat-tabs');
  container.innerHTML = strategyData.categories.map(cat =>
    `<div class="sub-tab${cat === strategyCat ? ' active' : ''}" onclick="setStrategyCat('${cat}')">${cat}</div>`
  ).join('');
}

function renderStrategySeriesChecks() {
  const container = document.getElementById('strategy-series-checks');
  const seriesList = strategyData.series_order[strategyCat] || [];
  container.innerHTML = seriesList.map((ser, i) => {
    const color   = STRATEGY_COLORS[i % STRATEGY_COLORS.length];
    const checked = selectedStrategySeries.has(ser);
    const escaped = ser.replace(/'/g, "\\'");
    return `<label class="week-check-label ${checked ? 'checked' : ''}" style="${checked ? 'background:' + color + ';border-color:' + color : ''}">
      <input type="checkbox" ${checked ? 'checked' : ''} onchange="toggleStrategySeries('${escaped}', this, '${color}', this.parentElement)">
      ${ser}
    </label>`;
  }).join('');
}

function toggleStrategySeries(ser, cb, color, label) {
  if (cb.checked) {
    selectedStrategySeries.add(ser);
    label.classList.add('checked');
    label.style.background  = color;
    label.style.borderColor = color;
  } else {
    selectedStrategySeries.delete(ser);
    label.classList.remove('checked');
    label.style.background  = '';
    label.style.borderColor = '';
  }
  renderStrategy();
}

function setStrategyCat(cat) {
  strategyCat = cat;
  selectedStrategySeries = new Set(strategyData.series_order[cat] || []);
  renderStrategyCatTabs();
  renderStrategySeriesChecks();
  renderStrategy();
}

function renderStrategy() {
  if (!strategyData || !strategyCat) return;
  const catData    = strategyData.data[strategyCat] || {};
  const seriesList = strategyData.series_order[strategyCat] || [];
  const weeks      = strategyData.weeks;

  const activeSeries = seriesList.filter(ser => selectedStrategySeries.has(ser));

  // 합계 카드 — 선택된 시리즈 합산
  let totalOffline = 0, totalOnline = 0, totalBiz = 0;
  activeSeries.forEach(ser => {
    weeks.forEach(w => {
      totalOffline += catData[ser]?.[w]?.offline || 0;
      totalOnline  += catData[ser]?.[w]?.online  || 0;
      totalBiz     += catData[ser]?.[w]?.biz     || 0;
    });
  });
  document.getElementById('strategy-totals').innerHTML = `
    <div class="total-card" style="border-left-color:#1a237e">
      <div class="wname">오프라인 합계</div>
      <div class="wamt" style="color:#1a237e">${totalOffline.toLocaleString('ko-KR')}</div>
    </div>
    <div class="total-card" style="border-left-color:#2e7d32">
      <div class="wname">온라인 합계</div>
      <div class="wamt" style="color:#2e7d32">${totalOnline.toLocaleString('ko-KR')}</div>
    </div>
    <div class="total-card" style="border-left-color:#e53935">
      <div class="wname">사업개발 합계</div>
      <div class="wamt" style="color:#e53935">${totalBiz.toLocaleString('ko-KR')}</div>
    </div>
  `;

  // X축=주차, 시리즈×채널별 라인
  // 색상=시리즈, 선스타일=채널(실선/파선/점선)
  const datasets = [];

  // 전체 합계 라인 (선택된 시리즈 × 전 채널 총합)
  const totalVals = weeks.map(w =>
    activeSeries.reduce((s, ser) =>
      s + (selectedStrategyChannels.has('offline') ? (catData[ser]?.[w]?.offline || 0) : 0)
        + (selectedStrategyChannels.has('online')  ? (catData[ser]?.[w]?.online  || 0) : 0)
        + (selectedStrategyChannels.has('biz')     ? (catData[ser]?.[w]?.biz     || 0) : 0), 0)
  );
  datasets.push({
    label: '전체 합계',
    data: totalVals,
    borderColor: '#ff6f00',
    backgroundColor: '#ff6f00',
    borderWidth: 3,
    pointBackgroundColor: '#ff6f00',
    pointBorderColor: '#fff',
    pointBorderWidth: 2,
    pointRadius: 9,
    pointHoverRadius: 12,
    pointStyle: 'star',
    tension: 0.3,
    order: -1,
    datalabels: { display: false }
  });

  activeSeries.forEach((ser, si) => {
    const color = STRATEGY_COLORS[seriesList.indexOf(ser) % STRATEGY_COLORS.length];
    const serData = catData[ser] || {};
    ['offline', 'online', 'biz'].filter(ch => selectedStrategyChannels.has(ch)).forEach(ch => {
      const vals = weeks.map(w => serData[w]?.[ch] ?? null);
      datasets.push({
        label: `${ser} (${CHANNEL_LABEL[ch]})`,
        data: vals,
        borderColor: color,
        backgroundColor: color,
        borderDash: CHANNEL_DASH[ch],
        pointBackgroundColor: color,
        pointBorderColor: '#fff',
        pointBorderWidth: 2,
        pointRadius: 7,
        pointHoverRadius: 10,
        pointStyle: CHANNEL_POINT[ch],
        tension: 0.3,
        datalabels: { display: false }
      });
    });
  });

  if (strategyChart) strategyChart.destroy();
  strategyChart = new Chart(document.getElementById('strategy-chart').getContext('2d'), {
    type: 'line',
    data: { labels: weeks, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'top',
          labels: { usePointStyle: true, pointStyleWidth: 20 }
        },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              const v = context.parsed.y;
              return ` ${context.dataset.label}: ${v !== null ? v.toLocaleString('ko-KR') : '-'}`;
            }
          }
        }
      },
      scales: {
        y: {
          title: { display: true, text: '수량' },
          ticks: { callback: v => v.toLocaleString('ko-KR') },
          min: 0
        },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

/* ===== 카테고리실적 ===== */
function renderCatPerfInit() {
  if (!strategyData || !strategyData.categories || !strategyData.categories.length) return;
  renderCategoryPerf();
}

function linRegression(vals) {
  const n = vals.length;
  const xs = vals.map((_, i) => i);
  const sumX  = xs.reduce((a, b) => a + b, 0);
  const sumY  = vals.reduce((a, b) => a + b, 0);
  const sumXY = xs.reduce((a, i) => a + i * vals[i], 0);
  const sumX2 = xs.reduce((a, i) => a + i * i, 0);
  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;
  return xs.map(i => Math.max(0, slope * i + intercept));
}

function renderCategoryPerf() {
  if (!strategyData || !strategyData.categories) return;
  const weeks = strategyData.weeks;

  const datasets = [];
  strategyData.categories.forEach((cat, ci) => {
    const color   = STRATEGY_COLORS[ci % STRATEGY_COLORS.length];
    const catData = strategyData.data[cat] || {};
    const serList = strategyData.series_order[cat] || [];
    const vals    = weeks.map(w =>
      serList.reduce((s, ser) =>
        s + (catData[ser]?.[w]?.offline || 0)
          + (catData[ser]?.[w]?.online  || 0)
          + (catData[ser]?.[w]?.biz     || 0), 0)
    );
    datasets.push({
      label: cat, data: vals,
      borderColor: color, backgroundColor: color,
      borderWidth: 2,
      pointBackgroundColor: color, pointBorderColor: '#fff', pointBorderWidth: 2,
      pointRadius: 7, pointHoverRadius: 10,
      tension: 0.3,
      datalabels: { display: false }
    });
    datasets.push({
      label: cat + ' 추세',
      data: linRegression(vals),
      borderColor: color + '4D', backgroundColor: 'transparent',
      borderWidth: 1.5, borderDash: [6, 4],
      pointRadius: 0, pointHoverRadius: 0,
      tension: 0,
      datalabels: { display: false }
    });
  });

  if (catPerfChart) catPerfChart.destroy();
  catPerfChart = new Chart(document.getElementById('cat-perf-chart').getContext('2d'), {
    type: 'line',
    data: { labels: weeks, datasets },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        datalabels: { display: true },
        tooltip: {
          callbacks: {
            label(context) {
              if (context.dataset.label.endsWith(' 추세')) return null;
              const v = context.parsed.y;
              return ` ${context.dataset.label}: ${v !== null ? v.toLocaleString('ko-KR') : '-'}`;
            }
          }
        }
      },
      scales: {
        y: {
          title: { display: true, text: '합계' },
          ticks: { callback: v => v.toLocaleString('ko-KR') },
          min: 0
        },
        x: { ticks: { maxRotation: 45 } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

loadAll();
</script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/api/series")
def api_series():
    return jsonify(fetch_series())

@app.route("/api/series_naver")
def api_series_naver():
    return jsonify(fetch_series_naver())

@app.route("/api/cvr")
def api_cvr():
    return jsonify(fetch_cvr())

@app.route("/api/alrim")
def api_alrim():
    return jsonify(fetch_alrim())

@app.route("/api/customer")
def api_customer():
    return jsonify(fetch_customer())

@app.route("/api/marketing")
def api_marketing():
    return jsonify(fetch_marketing())

@app.route("/api/nps")
def api_nps():
    return jsonify(fetch_nps())

@app.route("/api/strategy")
def api_strategy():
    return jsonify(fetch_strategy())

if __name__ == "__main__":
    app.run(debug=True, port=5001)
