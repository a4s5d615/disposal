#!/usr/bin/env python3
"""
處置中分盤交易清單 + 月線乖離率
=============================
每日執行，生成 disposal_ma20.html

用法：
  python disposal_ma20.py

依賴：requests, beautifulsoup4, lxml
"""

import re
import sys
import io
import time
import json
import urllib3
import requests
from datetime import datetime, timedelta
from pathlib import Path
from bs4 import BeautifulSoup

# 強制 stdout/stderr 使用 UTF-8（Windows cmd 預設為 cp950）
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# 停用 SSL 憑證驗證（Windows 上 TWSE/TPEX 的 SKI 問題）
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── 常數 ──────────────────────────────────────────────────────────────────────

OUTPUT = Path(__file__).parent / "disposal_ma20.html"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9",
}

REQUEST_DELAY = 0.5   # 每次 API 請求間隔（秒）


# ── HTTP 工具 ─────────────────────────────────────────────────────────────────

def fetch(url, params=None, retries=3, timeout=20):
    for i in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, params=params,
                             timeout=timeout, verify=False)
            r.encoding = "utf-8"
            if r.status_code == 200:
                return r
        except Exception as e:
            print(f"  [retry {i+1}] {url.split('?')[0]}: {type(e).__name__}", file=sys.stderr)
            if i < retries - 1:
                time.sleep(2 ** i)
    return None


# ── 爬取處置清單 ──────────────────────────────────────────────────────────────

def scrape_disposal_stocks():
    """
    爬取 chengwaye.com/disposal-forecast
    回傳 list of dict：market, code, name, interval, start, end, exit_date, remaining, reason
    """
    print("▶ 爬取處置清單 ...", flush=True)
    r = fetch("https://chengwaye.com/disposal-forecast")
    if not r:
        print("  ✗ 無法取得頁面", file=sys.stderr)
        return []

    soup = BeautifulSoup(r.text, "lxml")
    stocks = []

    # 找到標題含「處置中」的 section，再取其下的 <table>
    target_table = None
    for el in soup.find_all(string=re.compile(r"目前處置中")):
        parent = el.find_parent()
        if parent:
            tbl = parent.find_next("table")
            if tbl:
                target_table = tbl
                break

    # fallback：找第一個 thead 含「撮合」的表格
    if not target_table:
        for tbl in soup.find_all("table"):
            header_text = tbl.get_text()
            if "撮合" in header_text and "出關" in header_text:
                target_table = tbl
                break

    if not target_table:
        print("  ✗ 找不到處置表格", file=sys.stderr)
        return []

    for tr in target_table.find_all("tr"):
        # 優先使用 data-code 屬性
        code = tr.get("data-code", "").strip()
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]

        if not code:
            # 從 cells 中找 4 位數代號
            for c in cells:
                if re.match(r"^\d{4,5}$", c):
                    code = c
                    break

        if not code or not re.match(r"^\d{4,5}$", code):
            continue

        if len(cells) < 7:
            continue

        # 欄位順序：所, 代號, 名稱, 撮合, 開始, 結束, 出關日, 剩餘交易日, 處置原因
        try:
            # 找 code 在 cells 中的位置
            code_idx = next(
                (i for i, c in enumerate(cells) if c == code), 1
            )
            offset = code_idx - 1   # '所' 欄

            def cell(i):
                idx = offset + i
                return cells[idx] if 0 <= idx < len(cells) else ""

            stocks.append({
                "market":    cell(0),    # 市/櫃
                "code":      cell(1),
                "name":      cell(2),
                "interval":  cell(3),    # 5分/10分/20分
                "start":     cell(4),
                "end":       cell(5),
                "exit_date": cell(6),
                "remaining": cell(7),
                "reason":    cell(8),
                # 待填入
                "price":        None,
                "ma10":         None,
                "dev10":        None,
                "ma20":         None,
                "deviation":    None,
                "pre_close":    None,   # 封關日前一個交易日收盤
                "gain_from_pre": None,  # (現價 - pre_close) / pre_close × 100%
            })
        except Exception as e:
            print(f"  [parse] {e}", file=sys.stderr)

    print(f"  ✓ 共 {len(stocks)} 檔個股", flush=True)
    return stocks


# ── TWSE/TPEX 月收盤價 ────────────────────────────────────────────────────────

def twse_monthly_closes(code, year, month):
    """上市：TWSE STOCK_DAY 月資料，回傳 [close, ...]"""
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY"
    r = fetch(url, params={
        "response": "json",
        "date": f"{year}{month:02d}01",
        "stockNo": code,
    })
    if not r:
        return []
    try:
        data = r.json()
        if data.get("stat") != "OK":
            return []
        closes = []
        for row in data.get("data", []):
            try:
                closes.append(float(row[6].replace(",", "")))
            except Exception:
                pass
        return closes
    except Exception:
        return []


# ── 上櫃個股歷史收盤（Yahoo Finance） ────────────────────────────────────────

YF_BASE = "https://query1.finance.yahoo.com/v8/finance/chart"
YF_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
}


def yf_closes(code, suffix="TWO", days=35):
    """
    用 Yahoo Finance v8 API 取得上櫃個股近 N 天收盤。
    suffix='TWO' 上櫃；suffix='TW' 上市（備用）
    回傳 [close, ...] 由舊至新
    """
    end   = int(datetime.now().timestamp())
    start = int((datetime.now() - timedelta(days=days)).timestamp())
    url   = f"{YF_BASE}/{code}.{suffix}"
    try:
        r = requests.get(
            url,
            params={"period1": start, "period2": end, "interval": "1d"},
            headers=YF_HEADERS,
            timeout=15,
            verify=False,
        )
        if r.status_code != 200:
            return []
        result = r.json().get("chart", {}).get("result", [])
        if not result:
            return []
        closes_raw = result[0].get("indicators", {}).get("quote", [{}])[0].get("close", [])
        # 過濾 None 值，並四捨五入至 2 位小數
        return [round(float(c), 2) for c in closes_raw if c is not None]
    except Exception:
        return []


def collect_closes(code, market):
    """取得近 20+ 個交易日收盤價。"""
    if market != "市":
        closes = yf_closes(code, suffix="TWO", days=35)
        time.sleep(REQUEST_DELAY)
        return closes

    # TWSE 上市：月資料（不足 20 日補前月）
    now = datetime.now()
    y, m = now.year, now.month
    closes = twse_monthly_closes(code, y, m)
    time.sleep(REQUEST_DELAY)
    if len(closes) < 20:
        pm = m - 1 if m > 1 else 12
        py = y if m > 1 else y - 1
        prev = twse_monthly_closes(code, py, pm)
        time.sleep(REQUEST_DELAY)
        closes = prev + closes
    return closes


# ── 封關日前一個交易日收盤 ───────────────────────────────────────────────────────

def twse_closes_for_month(code, year, month):
    """上市：TWSE STOCK_DAY 月資料，回傳 [(date_str, close), ...]，date_str='YYYY-MM-DD'"""
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY"
    r = fetch(url, params={
        "response": "json",
        "date": f"{year}{month:02d}01",
        "stockNo": code,
    })
    if not r:
        return []
    try:
        data = r.json()
        if data.get("stat") != "OK":
            return []
        result = []
        for row in data.get("data", []):
            try:
                # row[0] = "115/04/30"（民國年）
                parts = row[0].split("/")
                gre_year = int(parts[0]) + 1911
                date_str = f"{gre_year}-{parts[1]}-{parts[2]}"
                close = float(row[6].replace(",", ""))
                result.append((date_str, close))
            except Exception:
                pass
        return result
    except Exception:
        return []


def yf_closes_range(code, suffix, range_start, range_end):
    """Yahoo Finance: 取得指定區間的 [(date_str, close), ...]，date_str='YYYY-MM-DD'"""
    p1 = int(range_start.timestamp())
    p2 = int(range_end.timestamp())
    url = f"{YF_BASE}/{code}.{suffix}"
    try:
        r = requests.get(
            url,
            params={"period1": p1, "period2": p2, "interval": "1d"},
            headers=YF_HEADERS,
            timeout=15,
            verify=False,
        )
        if r.status_code != 200:
            return []
        result = r.json().get("chart", {}).get("result", [])
        if not result:
            return []
        timestamps = result[0].get("timestamp", [])
        closes_raw = result[0].get("indicators", {}).get("quote", [{}])[0].get("close", [])
        pairs = []
        for ts, c in zip(timestamps, closes_raw):
            if c is None:
                continue
            # Yahoo Finance 時間戳是 UTC，台股以亞洲收盤時間對應日期
            date_str = datetime.utcfromtimestamp(ts).strftime("%Y-%m-%d")
            pairs.append((date_str, round(float(c), 2)))
        return pairs
    except Exception:
        return []


def pre_disposal_close(code, market, start_mmdd):
    """
    取得封關日前一個交易日的收盤價。
    start_mmdd: 來自 chengwaye 的「開始」欄，格式 "MM-DD"
    回傳 float 或 None
    """
    now = datetime.now()
    try:
        mm, dd = map(int, start_mmdd.split("-"))
    except Exception:
        return None

    # 推算完整年份：若 MM-DD 晚於今天，則是去年
    year = now.year
    try:
        start_dt = datetime(year, mm, dd)
    except ValueError:
        return None
    if start_dt > now:
        try:
            start_dt = datetime(year - 1, mm, dd)
        except ValueError:
            return None

    target = start_dt.strftime("%Y-%m-%d")

    if market == "市":
        # TWSE：取當月資料
        y, m = start_dt.year, start_dt.month
        rows = twse_closes_for_month(code, y, m)
        time.sleep(REQUEST_DELAY)

        # 找 target 之前最近的交易日
        for date_str, close in reversed(rows):
            if date_str < target:
                return close

        # 若 start_dt 在月初，需補前一個月
        pm = m - 1 if m > 1 else 12
        py = y if m > 1 else y - 1
        prev_rows = twse_closes_for_month(code, py, pm)
        time.sleep(REQUEST_DELAY)
        for date_str, close in reversed(prev_rows):
            if date_str < target:
                return close
        return None

    else:
        # 上櫃：Yahoo Finance 取前後區間
        range_start = start_dt - timedelta(days=10)
        range_end   = start_dt + timedelta(days=1)
        pairs = yf_closes_range(code, "TWO", range_start, range_end)
        time.sleep(REQUEST_DELAY)
        for date_str, close in reversed(pairs):
            if date_str < target:
                return close
        return None


# ── 即時股價 ──────────────────────────────────────────────────────────────────

def realtime_price(code, market):
    """
    取得即時（或最近）收盤價。
    上市：TWSE MIS API；上櫃：Yahoo Finance 最後一筆收盤。
    """
    if market != "市":
        # Yahoo Finance 最後收盤即為最新價
        # （collect_closes 已取過，這裡只補 fallback 用）
        closes = yf_closes(code, suffix="TWO", days=5)
        return closes[-1] if closes else None

    r = fetch(
        "https://mis.twse.com.tw/stock/api/getStockInfo.jsp",
        params={"ex_ch": f"tse_{code}.tw", "json": "1", "delay": "0"},
    )
    if not r:
        return None
    try:
        msg = r.json().get("msgArray", [{}])[0]
        z = msg.get("z", "-")
        if z and z not in ("-", ""):
            return float(z)
        y_val = msg.get("y", "")
        if y_val and y_val not in ("-", ""):
            return float(y_val)
    except Exception:
        pass
    return None


# ── 計算乖離率 ────────────────────────────────────────────────────────────────

def enrich_stock(stock):
    """
    為個股填入：price（現價）, ma20（月線）, deviation（月線乖離率 %）
    """
    code   = stock["code"]
    market = stock["market"]

    print(f"  {code} {stock['name']} ({market}) ...", end=" ", flush=True)

    closes = collect_closes(code, market)
    if not closes:
        print("無收盤資料", flush=True)
        return

    if market != "市":
        # 上櫃：Yahoo Finance 最後一筆即為當前價，無需再打一次 API
        price = closes[-1]
    else:
        price = realtime_price(code, market)
        time.sleep(REQUEST_DELAY)
        if price is None:
            price = closes[-1]

    # MA10（十日線）
    w10   = closes[-10:] if len(closes) >= 10 else closes
    ma10  = round(sum(w10) / len(w10), 2)
    dev10 = round((price - ma10) / ma10 * 100, 2)

    # MA20（月線）
    w20   = closes[-20:] if len(closes) >= 20 else closes
    ma20  = round(sum(w20) / len(w20), 2)
    dev20 = round((price - ma20) / ma20 * 100, 2)

    # 封關日前一個交易日收盤
    pre_close = pre_disposal_close(code, market, stock.get("start", ""))
    gain_from_pre = (
        round((price - pre_close) / pre_close * 100, 2)
        if pre_close and pre_close > 0
        else None
    )

    stock["price"]        = price
    stock["ma10"]         = ma10
    stock["dev10"]        = dev10
    stock["ma20"]         = ma20
    stock["deviation"]    = dev20
    stock["pre_close"]    = pre_close
    stock["gain_from_pre"] = gain_from_pre

    pre_str = f"封關前={pre_close}({gain_from_pre:+.2f}%)  " if gain_from_pre is not None else ""
    print(
        f"現價={price}  {pre_str}MA10={ma10}({dev10:+.2f}%)  MA20={ma20}({dev20:+.2f}%)",
        flush=True
    )


# ── HTML 生成 ─────────────────────────────────────────────────────────────────

HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>處置中分盤交易 — 月線乖離率 {date}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#0d1117;color:#e6edf3;font-family:-apple-system,'Segoe UI',sans-serif;font-size:14px;line-height:1.6}}
a{{color:#58a6ff;text-decoration:none}}

nav{{background:#0a0a16;border-bottom:1px solid #21262d;padding:12px 24px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;position:sticky;top:0;z-index:100}}
nav h1{{font-size:17px;font-weight:700}}
.badge{{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}}
.badge-blue{{background:#1c2e4a;color:#58a6ff}}
.badge-gray{{background:#21262d;color:#8b949e}}
.badge-yellow{{background:#3a2d00;color:#d29922}}
.update{{margin-left:auto;font-size:11px;color:#8b949e}}

.container{{max-width:1200px;margin:24px auto;padding:0 16px}}

.summary{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}}
.kpi{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:14px 20px;text-align:center;min-width:110px}}
.kpi .num{{font-size:26px;font-weight:700;color:#58a6ff}}
.kpi .lbl{{font-size:11px;color:#8b949e;margin-top:2px}}

.card{{background:#161b22;border:1px solid #21262d;border-radius:10px;overflow:hidden}}
.card-header{{padding:14px 16px;border-bottom:1px solid #21262d;display:flex;align-items:center;gap:10px}}
.card-header h2{{font-size:14px;font-weight:600;color:#8b949e;text-transform:uppercase;letter-spacing:.04em}}

table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{text-align:left;padding:9px 12px;color:#8b949e;font-weight:600;font-size:11px;text-transform:uppercase;border-bottom:1px solid #21262d;white-space:nowrap;cursor:pointer;user-select:none}}
th:hover{{color:#e6edf3}}
th.sort-asc::after{{content:" ↑"}}
th.sort-desc::after{{content:" ↓"}}
td{{padding:8px 12px;border-bottom:1px solid #0d1117;white-space:nowrap}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:#1c2128}}
.num-cell{{text-align:right;font-variant-numeric:tabular-nums;font-family:'SF Mono','Fira Code',monospace}}
.code-cell{{font-family:monospace;color:#58a6ff;font-weight:600}}
.code-cell a{{color:inherit;text-decoration:none}}
.code-cell a:hover{{text-decoration:underline}}
.name-cell{{font-weight:500}}
.name-cell a{{color:inherit;text-decoration:none}}
.name-cell a:hover{{text-decoration:underline;color:#58a6ff}}
.positive{{color:#f85149}}   /* 乖離率為正 = 超過月線 = 偏高（紅） */
.negative{{color:#3fb950}}   /* 乖離率為負 = 低於月線 = 偏低（綠） */
.neutral{{color:#8b949e}}
.no-data{{color:#484f58;font-style:italic}}

.dev-bar{{display:flex;align-items:center;gap:8px}}
.bar-track{{width:80px;height:6px;background:#21262d;border-radius:3px;overflow:hidden;flex-shrink:0}}
.bar-pos{{height:100%;background:#f85149;border-radius:3px}}
.bar-neg{{height:100%;background:#3fb950;border-radius:3px;margin-left:auto}}

.reason-cell{{max-width:200px;overflow:hidden;text-overflow:ellipsis;color:#8b949e;font-size:12px}}
.interval-badge{{display:inline-block;padding:1px 6px;border-radius:4px;font-size:11px;background:#21262d;color:#8b949e}}
.remaining-badge{{display:inline-block;padding:1px 6px;border-radius:4px;font-size:11px;font-weight:700}}
.rem-exit{{background:#1a4731;color:#3fb950}}
.rem-soon{{background:#3a2d00;color:#d29922}}
.rem-long{{background:#1c2e4a;color:#58a6ff}}
.market-m{{color:#f85149;font-weight:700}}
.market-k{{color:#58a6ff;font-weight:700}}

.search-box{{padding:8px 12px;background:#0d1117;border:1px solid #21262d;border-radius:6px;color:#e6edf3;font-size:13px;width:220px;outline:none}}
.search-box:focus{{border-color:#58a6ff}}
.filter-row{{display:flex;gap:8px;flex-wrap:wrap;padding:12px 16px;border-bottom:1px solid #21262d;align-items:center}}
.filter-btn{{padding:4px 12px;border-radius:6px;border:1px solid #21262d;background:#0d1117;color:#8b949e;font-size:12px;cursor:pointer}}
.filter-btn.active{{background:#1c2e4a;color:#58a6ff;border-color:#58a6ff}}

footer{{text-align:center;color:#484f58;font-size:11px;padding:24px;margin-top:24px;border-top:1px solid #21262d}}
</style>
</head>
<body>

<nav>
  <h1>⚫ 處置中分盤交易</h1>
  <span class="badge badge-blue">資料來源：chengwaye.com</span>
  <span class="badge badge-yellow">月線乖離率</span>
  <span class="update">基準日 {date} ｜ 共 {total} 檔 ｜ 產生時間 {now}</span>
</nav>

<div class="container">

  <!-- KPI -->
  <div class="summary">
    <div class="kpi">
      <div class="num">{total}</div>
      <div class="lbl">分盤交易總數</div>
    </div>
    <div class="kpi">
      <div class="num" style="color:#3fb950">{cnt_exit}</div>
      <div class="lbl">明日出關</div>
    </div>
    <div class="kpi">
      <div class="num" style="color:#f85149">{cnt_pos_dev}</div>
      <div class="lbl">乖離率 &gt;+5%</div>
    </div>
    <div class="kpi">
      <div class="num" style="color:#3fb950">{cnt_neg_dev}</div>
      <div class="lbl">乖離率 &lt;-5%</div>
    </div>
    <div class="kpi">
      <div class="num" style="color:#8b949e">{cnt_no_data}</div>
      <div class="lbl">無法取得資料</div>
    </div>
  </div>

  <!-- 主表格 -->
  <div class="card">
    <div class="card-header">
      <h2>個股清單</h2>
      <input class="search-box" id="searchBox" type="text" placeholder="搜尋代號或名稱...">
    </div>
    <div class="filter-row">
      <span style="font-size:12px;color:#8b949e">篩選：</span>
      <button class="filter-btn active" data-filter="all">全部</button>
      <button class="filter-btn" data-filter="exit">明日出關</button>
      <button class="filter-btn" data-filter="pos">乖離率 &gt;0%</button>
      <button class="filter-btn" data-filter="neg">乖離率 &lt;0%</button>
      <button class="filter-btn" data-filter="market">上市</button>
      <button class="filter-btn" data-filter="otc">上櫃</button>
    </div>
    <div style="overflow-x:auto">
    <table id="mainTable">
      <thead>
        <tr>
          <th data-col="0">所</th>
          <th data-col="1">代號</th>
          <th data-col="2">名稱</th>
          <th data-col="3">撮合</th>
          <th data-col="4">開始</th>
          <th data-col="5">結束</th>
          <th data-col="6">出關日</th>
          <th data-col="7" class="sort-asc">剩餘</th>
          <th data-col="8" class="num-cell">現價</th>
          <th data-col="9" class="num-cell" title="(現價 - 封關日前一個交易日收盤) / 封關前收盤 × 100%">封關漲幅</th>
          <th data-col="10" class="num-cell">MA10</th>
          <th data-col="11" class="num-cell">十日線乖離率</th>
          <th data-col="12" class="num-cell">MA20</th>
          <th data-col="13" class="num-cell">月線乖離率</th>
          <th>處置原因</th>
        </tr>
      </thead>
      <tbody id="tableBody">
{rows}
      </tbody>
    </table>
    </div>
  </div>

  <div style="margin-top:12px;font-size:11px;color:#484f58;line-height:2">
    ⓘ 封關漲幅 = (現價 − 封關日前一個交易日收盤) ÷ 封關前收盤 × 100%<br>
    ⓘ 十日線乖離率 = (現價 − MA10) ÷ MA10 × 100%　|　MA10 = 近10個交易日收盤均值<br>
    ⓘ 月線乖離率 = (現價 − MA20) ÷ MA20 × 100%　|　MA20 = 近20個交易日收盤均值<br>
    ⓘ 紅色 = 高於基準，綠色 = 低於基準　|　現價：上市用 TWSE MIS，上櫃用 Yahoo Finance<br>
    ⚠️ 本頁資訊僅供研究參考，不構成投資建議
  </div>

</div>

<footer>
  資料來源：chengwaye.com / TWSE / TPEX ｜ 本頁資訊僅供研究參考，不構成投資建議
</footer>

<script>
// ── 排序 ──────────────────────────────────────────────────────────────────────
const table  = document.getElementById('mainTable');
const tbody  = document.getElementById('tableBody');
let sortCol  = 7;
let sortAsc  = true;
let allRows  = [];

function sortTable(col) {{
  if (sortCol === col) sortAsc = !sortAsc;
  else {{ sortCol = col; sortAsc = true; }}
  table.querySelectorAll('th').forEach((th, i) => {{
    th.classList.remove('sort-asc','sort-desc');
    if (i === col) th.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');
  }});
  applyFilterAndSort();
}}

function parseVal(td, col) {{
  const text = td.dataset.val ?? td.textContent.trim();
  if (col >= 8) return parseFloat(text) || -Infinity;
  if (col === 7) {{
    if (text === '出關') return -1;
    return parseInt(text) || 999;
  }}
  return text;
}}

table.querySelectorAll('th[data-col]').forEach(th => {{
  th.addEventListener('click', () => sortTable(parseInt(th.dataset.col)));
}});

// ── 搜尋 + 篩選 ───────────────────────────────────────────────────────────────
let currentFilter = 'all';

document.getElementById('searchBox').addEventListener('input', applyFilterAndSort);
document.querySelectorAll('.filter-btn').forEach(btn => {{
  btn.addEventListener('click', () => {{
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    applyFilterAndSort();
  }});
}});

function applyFilterAndSort() {{
  const q    = document.getElementById('searchBox').value.toLowerCase();
  const rows = Array.from(tbody.querySelectorAll('tr'));

  // 過濾
  rows.forEach(tr => {{
    const code  = (tr.dataset.code  || '').toLowerCase();
    const name  = (tr.dataset.name  || '').toLowerCase();
    const mkt   = tr.dataset.market || '';
    const rem   = tr.dataset.rem    || '';
    const dev   = parseFloat(tr.dataset.dev ?? 'NaN');

    let show = true;
    if (q && !code.includes(q) && !name.includes(q)) show = false;
    if (currentFilter === 'exit'   && rem !== '出關') show = false;
    if (currentFilter === 'pos'    && !(dev > 0))     show = false;
    if (currentFilter === 'neg'    && !(dev < 0))     show = false;
    if (currentFilter === 'market' && mkt !== '市')   show = false;
    if (currentFilter === 'otc'    && mkt !== '櫃')   show = false;

    tr.style.display = show ? '' : 'none';
  }});

  // 排序可見列
  const visible = rows.filter(tr => tr.style.display !== 'none');
  visible.sort((a, b) => {{
    const va = parseVal(a.cells[sortCol], sortCol);
    const vb = parseVal(b.cells[sortCol], sortCol);
    if (va < vb) return sortAsc ? -1 : 1;
    if (va > vb) return sortAsc ?  1 : -1;
    return 0;
  }});
  visible.forEach(tr => tbody.appendChild(tr));
}}
</script>
</body>
</html>
"""

ROW_TEMPLATE = """\
        <tr data-code="{code}" data-name="{name}" data-market="{market}"
            data-rem="{remaining}" data-dev="{dev_val}" data-dev10="{dev10_val}">
          <td class="{market_cls}">{market}</td>
          <td class="code-cell"><a href="https://goodinfo.tw/tw/StockInfo.asp?STOCK_ID={code}" target="_blank" rel="noopener">{code}</a></td>
          <td class="name-cell"><a href="https://tw.stock.yahoo.com/quote/{code}" target="_blank" rel="noopener">{name}</a></td>
          <td><span class="interval-badge">{interval}</span></td>
          <td>{start}</td>
          <td>{end}</td>
          <td>{exit_date}</td>
          <td><span class="remaining-badge {rem_cls}">{remaining}</span></td>
          <td class="num-cell" data-val="{price_raw}">{price_str}</td>
          <td class="num-cell" data-val="{gain_raw}">{gain_cell}</td>
          <td class="num-cell" data-val="{ma10_raw}">{ma10_str}</td>
          <td class="num-cell">{dev10_cell}</td>
          <td class="num-cell" data-val="{ma20_raw}">{ma20_str}</td>
          <td class="num-cell">{dev_cell}</td>
          <td class="reason-cell" title="{reason}">{reason_short}</td>
        </tr>"""


def remaining_class(rem):
    if rem == "出關":
        return "rem-exit"
    try:
        d = int(rem.replace("天", ""))
        return "rem-soon" if d <= 3 else "rem-long"
    except Exception:
        return "rem-long"


def deviation_cell(dev):
    if dev is None:
        return '<span class="no-data">—</span>'

    cls = "positive" if dev > 0 else "negative" if dev < 0 else "neutral"
    sign = "+" if dev > 0 else ""
    pct = f"{sign}{dev:.2f}%"

    # mini bar（最大顯示 ±20%）
    capped = min(abs(dev), 20) / 20 * 100
    if dev > 0:
        bar = f'<div class="bar-track"><div class="bar-pos" style="width:{capped:.0f}%"></div></div>'
    else:
        bar = f'<div class="bar-track"><div class="bar-neg" style="width:{capped:.0f}%"></div></div>'

    return f'<div class="dev-bar">{bar}<span class="{cls}">{pct}</span></div>'


def render_html(stocks):
    now   = datetime.now()
    date  = now.strftime("%Y-%m-%d")
    now_s = now.strftime("%Y-%m-%d %H:%M:%S")

    total     = len(stocks)
    cnt_exit  = sum(1 for s in stocks if s["remaining"] == "出關")
    cnt_pos   = sum(1 for s in stocks if s["deviation"] is not None and s["deviation"] > 5)
    cnt_neg   = sum(1 for s in stocks if s["deviation"] is not None and s["deviation"] < -5)
    cnt_nodat = sum(1 for s in stocks if s["deviation"] is None)

    rows_html = []
    for s in stocks:
        dev      = s["deviation"]
        dev10    = s["dev10"]
        price    = s["price"]
        ma10     = s["ma10"]
        ma20     = s["ma20"]
        gain     = s.get("gain_from_pre")
        dev_val  = f"{dev:.2f}"  if dev  is not None else "nan"
        dev10_val= f"{dev10:.2f}" if dev10 is not None else "nan"

        price_str = f"{price:.2f}" if price is not None else "—"
        ma10_str  = f"{ma10:.2f}"  if ma10  is not None else "—"
        ma20_str  = f"{ma20:.2f}"  if ma20  is not None else "—"
        reason    = s.get("reason", "")
        reason_sh = reason[:18] + "…" if len(reason) > 20 else reason

        rows_html.append(ROW_TEMPLATE.format(
            code        = s["code"],
            name        = s["name"],
            market      = s["market"],
            market_cls  = "market-m" if s["market"] == "市" else "market-k",
            interval    = s["interval"],
            start       = s["start"],
            end         = s["end"],
            exit_date   = s["exit_date"],
            remaining   = s["remaining"],
            rem_cls     = remaining_class(s["remaining"]),
            price_str   = price_str,
            price_raw   = f"{price:.2f}" if price is not None else "",
            gain_cell   = deviation_cell(gain),
            gain_raw    = f"{gain:.2f}" if gain is not None else "",
            ma10_str    = ma10_str,
            ma10_raw    = f"{ma10:.2f}" if ma10 is not None else "",
            dev10_cell  = deviation_cell(dev10),
            dev10_val   = dev10_val,
            ma20_str    = ma20_str,
            ma20_raw    = f"{ma20:.2f}" if ma20 is not None else "",
            dev_cell    = deviation_cell(dev),
            dev_val     = dev_val,
            reason      = reason,
            reason_short= reason_sh,
        ))

    html = HTML_TEMPLATE.format(
        date      = date,
        now       = now_s,
        total     = total,
        cnt_exit  = cnt_exit,
        cnt_pos_dev = cnt_pos,
        cnt_neg_dev = cnt_neg,
        cnt_no_data = cnt_nodat,
        rows      = "\n".join(rows_html),
    )
    return html


# ── 主程式 ────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  處置中分盤交易 + 月線乖離率")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    stocks = scrape_disposal_stocks()
    if not stocks:
        print("沒有取到資料，結束。")
        return

    print(f"\n▶ 取得 {len(stocks)} 檔個股的價格資料 ...\n", flush=True)
    for i, s in enumerate(stocks, 1):
        print(f"  [{i:2d}/{len(stocks)}]", end=" ")
        enrich_stock(s)

    print("\n▶ 生成 HTML ...", flush=True)
    html = render_html(stocks)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"  ✓ 已儲存至：{OUTPUT}")
    print("=" * 60)


if __name__ == "__main__":
    main()
