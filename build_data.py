#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
A股宽基ETF市值 / 指数总市值 比例 —— 数据构建脚本。

数据源：
  ETF 份额   深交所(ShowReport 范围查询) + 上交所(commonQuery.do 逐日)
  ETF 价格   新浪 CN_MarketData.getKLineData（宽基ETF价格≈净值）
  指数点位   新浪 CN_MarketData.getKLineData
  指数流通市值 东财 push2 (带重试，有兜底)

输出：index.html（自包含，内嵌数据 + ECharts 曲线）
支持增量：有缓存时只拉最近数据，秒级刷新
"""

import json
import os
import io
import time
import random
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache")
os.makedirs(CACHE, exist_ok=True)

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

START = "2023-01-01"
END = "2026-11-01"
TOP_N = 10  # 每个指数取前 N 大 ETF

# 指数：东财 secid + 新浪代码
INDEX_META = {
    "沪深300": {"secid": "1.000300", "sina": "sh000300"},
    "中证500": {"secid": "1.000905", "sina": "sh000905"},
    "上证50":  {"secid": "1.000016", "sina": "sh000016"},
    "创业板指": {"secid": "0.399006", "sina": "sz399006"},
    "科创50":  {"secid": "1.000688", "sina": "sh000688"},
}

# 指数流通市值兜底值（元），push2 不稳定时使用（2026-11-01 实测，部分为估算）
MCAP_FALLBACK = {
    "1.000300": 48.16e12,   # 沪深300 流通市值
    "1.000905": 12.36e12,   # 中证500（估，总市值×0.88）
    "1.000016": 21.61e12,   # 上证50（估，总市值×0.90）
    "0.399006": 12.80e12,   # 创业板指（估，总市值×0.88）
    "1.000688": 3.622e12,   # 科创50 流通市值
}

# 跟踪各指数的主要 ETF 候选（最终按最新规模取前 N 大）
ETF_CANDIDATES = {
    "沪深300": ["510300", "159919", "510330", "510310", "515330", "510350", "159925",
                "510390", "515380", "159673", "515660", "515310", "159300", "159330", "159393"],
    "中证500": ["510500", "159922", "512500", "512510", "510510", "159968", "159820",
                "159935", "159982", "159500", "159337", "510580"],
    "上证50":  ["510050", "510800", "510710", "510600", "510190"],
    "创业板指": ["159915", "159952", "159948", "159977", "159958", "159908", "159971",
                 "159964", "159956", "159957", "159821", "159810", "159205", "159247"],
    "科创50":  ["588000", "588080", "588050", "588090", "588180", "588020", "588450", "588870", "589360"],
}


def cache_path(name):
    return os.path.join(CACHE, name + ".json")


def load_cache(name):
    p = cache_path(name)
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return None


def save_cache(name, obj):
    with open(cache_path(name), "w", encoding="utf-8") as f:
        json.dump(obj, f)


def _merge(dst, src):
    """把 src 合并进 dst（src 里 dict 值按内层 dict 合并，其余覆盖）。"""
    for k, v in src.items():
        if isinstance(v, dict):
            dst.setdefault(k, {}).update(v)
        else:
            dst[k] = v
    return dst


RECENT_DAYS = 25  # 增量模式：每次只拉最近 N 天的新数据


# ============ ETF 份额 ============
def fetch_szse_shares(start: str, end: str):
    """深交所 ETF 每日份额（份）。返回 {code: {date: shares}}"""
    out = {}
    sd = pd.Timestamp(start)
    ed = pd.Timestamp(end)
    # 分 6 个月 chunk 拉取
    cur = sd
    while cur < ed:
        chunk_end = min(cur + pd.DateOffset(months=3) - pd.Timedelta(days=1), ed)
        s = cur.strftime("%Y%m%d")
        e = chunk_end.strftime("%Y%m%d")
        try:
            df = _szse_scale(s, e)
            for _, r in df.iterrows():
                code = r["基金代码"]
                if code not in out:
                    out[code] = {}
                out[code][str(r["日期"])[:10]] = float(r["基金份额"])
        except Exception as ex:
            print(f"  SZSE {s}~{e} 失败: {ex}")
        cur = chunk_end + pd.Timedelta(days=1)
        time.sleep(0.3)
    return out


def _szse_scale(start: str, end: str):
    url = "https://www.szse.cn/api/report/ShowReport"
    headers = {"Host": "www.szse.cn",
               "Referer": "https://www.szse.cn/market/fund/volume/etf/index.html",
               "User-Agent": UA["User-Agent"]}
    params = {"SHOWTYPE": "xlsx", "CATALOGID": "scsj_fund_jjgm", "TABKEY": "tab1",
              "txtStart": f"{start[:4]}-{start[4:6]}-{start[6:]}",
              "txtEnd": f"{end[:4]}-{end[4:6]}-{end[6:]}",
              "jjlb": "ETF", "random": str(random.random())}
    r = requests.get(url, params=params, headers=headers, timeout=30)
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content), engine="openpyxl").dropna(how="all")
    if df.empty:
        return pd.DataFrame(columns=["日期", "基金代码", "基金份额"])
    df = df.rename(columns={"基金规模(份)": "基金份额"})
    if "基金代码" in df.columns:
        df["基金代码"] = pd.to_numeric(df["基金代码"], errors="coerce")
        df = df[df["基金代码"].notna()]
        df["基金代码"] = df["基金代码"].astype(int).astype(str).str.zfill(6)
    if "日期" in df.columns:
        df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
        df = df[df["日期"].notna()]
    if "基金份额" in df.columns:
        df["基金份额"] = df["基金份额"].astype(str).str.replace(",", "", regex=False)
        df["基金份额"] = pd.to_numeric(df["基金份额"], errors="coerce")
    for c in ["日期", "基金代码", "基金份额"]:
        if c not in df.columns:
            df[c] = np.nan
    return df[["日期", "基金代码", "基金份额"]]


def fetch_sse_shares(dates):
    """上交所 ETF 每日份额（份）。dates 为交易日列表（YYYY-MM-DD）。返回 {code: {date: shares}}"""
    out = {}
    url = "https://query.sse.com.cn/commonQuery.do"
    headers = {"Referer": "https://www.sse.com.cn/", "User-Agent": UA["User-Agent"]}
    total = len(dates)
    for i, d in enumerate(dates):
        params = {"isPagination": "true", "pageHelp.pageSize": "10000",
                  "pageHelp.pageNo": "1", "pageHelp.beginPage": "1",
                  "pageHelp.cacheSize": "1", "pageHelp.endPage": "1",
                  "sqlId": "COMMON_SSE_ZQPZ_ETFZL_XXPL_ETFGM_SEARCH_L",
                  "STAT_DATE": d}
        try:
            r = requests.get(url, params=params, headers=headers, timeout=30)
            data = r.json()
            for row in data.get("result", []):
                code = str(row.get("SEC_CODE", ""))
                stat = row.get("STAT_DATE", "")
                vol = row.get("TOT_VOL")
                if code and vol is not None:
                    out.setdefault(code, {})[str(stat)[:10]] = float(vol) * 10000
        except Exception:
            pass
        if (i + 1) % 100 == 0:
            print(f"  SSE 进度 {i + 1}/{total}")
    return out


# ============ ETF 价格（≈净值，宽基ETF溢价极小） ============
def fetch_price(code, datalen=1023):
    """新浪 ETF 日线收盘价。宽基ETF价格≈净值（溢价<0.5%），
    且用户口径「市值」= 份额×价格。返回 {date: close}"""
    prefix = "sh" if code.startswith(("5", "6", "9")) else "sz"
    r = requests.get("https://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData",
                     params={"symbol": f"{prefix}{code}", "scale": "240", "ma": "no", "datalen": str(datalen)},
                     headers=UA, timeout=30)
    return {d["day"]: float(d["close"]) for d in r.json()}


# ============ 指数点位 & 总市值 ============
def fetch_index_level(sina_symbol, datalen=1023):
    """新浪指数日线收盘。返回 {date: close}"""
    r = requests.get("https://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData",
                     params={"symbol": sina_symbol, "scale": "240", "ma": "no", "datalen": str(datalen)},
                     headers=UA, timeout=30)
    return {d["day"]: float(d["close"]) for d in r.json()}


def fetch_index_mcap(secid):
    """东财指数流通市值（当前值，带重试）。"""
    for _ in range(6):
        try:
            r = requests.get("https://push2.eastmoney.com/api/qt/stock/get",
                             params={"secid": secid, "fields": "f117"},
                             headers={"User-Agent": UA["User-Agent"],
                                      "Referer": "https://quote.eastmoney.com/"},
                             timeout=8)
            d = r.json().get("data")
            if d and d.get("f117"):
                return float(d["f117"])
        except Exception:
            pass
        time.sleep(3)
    return None


# ============ 主流程 ============
def main():
    today = datetime.now()
    end_str = today.strftime("%Y-%m-%d")  # 动态结束日期（含今天）

    # 1. 深交所份额（有缓存则增量，否则全量）
    szse = load_cache("szse_shares") or {}
    if not szse:
        print("首次拉取深交所份额（全量）...")
        szse = fetch_szse_shares(START, end_str)
    else:
        recent_start = (today - timedelta(days=RECENT_DAYS)).strftime("%Y-%m-%d")
        _merge(szse, fetch_szse_shares(recent_start, end_str))
    save_cache("szse_shares", szse)
    print(f"深交所份额: {len(szse)} 只基金")

    # 2. 指数点位（增量）
    levels = {}
    for name, meta in INDEX_META.items():
        lv = load_cache(f"level_{name}")
        if not lv:
            lv = fetch_index_level(meta["sina"])
        else:
            lv.update(fetch_index_level(meta["sina"], RECENT_DAYS))
        save_cache(f"level_{name}", lv)
        levels[name] = lv
        print(f"指数点位 {name}: {len(lv)} 日")

    trade_dates = sorted(levels["沪深300"].keys())
    trade_dates = [d for d in trade_dates if d >= START]

    # 3. 上交所份额（有缓存则只拉最近10个交易日，否则全量）
    sse = load_cache("sse_shares") or {}
    if not sse:
        print(f"首次拉取上交所份额（{len(trade_dates)} 个交易日，较慢）...")
        sse = fetch_sse_shares(trade_dates)
    else:
        _merge(sse, fetch_sse_shares(trade_dates[-10:]))
    save_cache("sse_shares", sse)
    print(f"上交所份额: {len(sse)} 只基金")

    # 4. 各指数：市值 + 比例
    result = {}
    for name, meta in INDEX_META.items():
        candidates = ETF_CANDIDATES[name]
        # 价格（增量）
        prices = {}
        for code in candidates:
            px = load_cache(f"price_{code}")
            if not px:
                px = fetch_price(code)
            else:
                px.update(fetch_price(code, RECENT_DAYS))
            save_cache(f"price_{code}", px)
            if px:
                prices[code] = px
        # 市值 = 份额 × 价格
        etf_aum = {}  # {code: {date: 市值}}
        for code in candidates:
            shares = szse.get(code) or sse.get(code) or {}
            px = prices.get(code) or {}
            series = {}
            for d, sh in shares.items():
                if d in px:
                    series[d] = sh * px[d]
            if series:
                etf_aum[code] = series
        # 按最新 AUM 排序，取前 N
        latest_aum = {c: max(s.values()) for c, s in etf_aum.items() if s}
        top_codes = sorted(latest_aum, key=latest_aum.get, reverse=True)[:TOP_N]
        # 汇总前 N 大的 AUM / 份额（所有日期，历史日期 ETF 未上市时自然缺失是正常的）
        total_aum = {}
        for code in top_codes:
            for d, v in etf_aum.get(code, {}).items():
                total_aum[d] = total_aum.get(d, 0.0) + v
        total_shares = {}
        for code in top_codes:
            shares_map = szse.get(code) or sse.get(code) or {}
            for d, sh in shares_map.items():
                total_shares[d] = total_shares.get(d, 0.0) + sh
        # 只剔除「尾部」不完整日期（深市份额 T+1 晚发布，只影响最新 1-2 天）：
        # 从最新日期往回，剔除到第一个「所有 top_codes 都有份额」的完整日期为止
        for d in sorted(total_shares.keys(), reverse=True):
            if all((szse.get(c) or sse.get(c) or {}).get(d) is not None for c in top_codes):
                break
            total_aum.pop(d, None)
            total_shares.pop(d, None)
        # 指数总市值历史 = 当前总市值 × (点位/最新点位)
        mcap_now = fetch_index_mcap(meta["secid"])
        if mcap_now is None:
            mcap_now = MCAP_FALLBACK.get(meta["secid"])
        lv = levels[name]
        latest_level = max(lv.values())
        mcap_series = {}
        for d, level in lv.items():
            if mcap_now:
                mcap_series[d] = mcap_now * (level / latest_level)
        # 比例 = ETF总AUM / 指数总市值
        ratio = {}
        for d in total_aum:
            if d in mcap_series and mcap_series[d] > 0:
                ratio[d] = total_aum[d] / mcap_series[d] * 100  # 百分比
        result[name] = {
            "top_codes": top_codes,
            "etf_aum_total": total_aum,   # 元
            "total_shares": total_shares, # 份
            "mcap": mcap_series,          # 元
            "ratio": ratio,               # %
            "mcap_now": mcap_now,
            "level": levels[name],        # 指数点位
        }
        latest_ratio = ratio[max(ratio)] if ratio else None
        print(f"{name}: 前{TOP_N}大 ETF {top_codes}, 比例当前 {latest_ratio:.3f}%, 峰值 {max(ratio.values()):.3f}%" if ratio else f"{name}: 无比例数据")

    save_cache("result", result)

    # 3. 生成 HTML
    html = render_html(result)
    out_html = os.path.join(HERE, "index.html")
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n已生成 {out_html}")

    # 4. 微信推送（可选，SERVER_KEY 环境变量存在时）
    sendkey = os.environ.get("SERVER_KEY")
    if sendkey:
        lines = []
        for name in result:
            ratio_map = result[name]["ratio"]
            shares_map = result[name].get("total_shares", {})
            level_map = result[name].get("level", {})
            if not ratio_map:
                continue
            ds = sorted(ratio_map)
            cur_d, prev_d = ds[-1], (ds[-2] if len(ds) >= 2 else None)
            seg = f"占比 {ratio_map[cur_d]:.2f}%"
            if prev_d:
                seg += f"（{ratio_map[cur_d] - ratio_map[prev_d]:+.2f}pp）"
            if cur_d in shares_map:
                seg += f"；份额 {shares_map[cur_d] / 1e8:.1f}亿份"
                if prev_d in shares_map:
                    seg += f"（{(shares_map[cur_d] - shares_map[prev_d]) / 1e8:+.1f}亿份）"
            if cur_d in level_map:
                seg += f"；指数点位 {level_map[cur_d]:.1f}"
            lines.append(f"- {name} · {seg}")
        title = f"📊 ETF占比更新 {today.strftime('%Y-%m-%d')}"
        content = "前10大ETF市值 / 指数流通市值\n\n" + "\n".join(lines)
        try:
            r = requests.post(f"https://sctapi.ftqq.com/{sendkey}.send",
                              data={"title": title, "desp": content}, timeout=15)
            print(f"微信推送: code={r.json().get('code')}")
        except Exception as e:
            print(f"微信推送失败: {e}")


def render_html(result):
    """生成自包含 HTML（内嵌数据 + ECharts，每指数一张双轴图）。"""
    all_dates = set()
    for r in result.values():
        all_dates.update(r["ratio"].keys())
    all_dates = sorted(all_dates)

    payload = {"dates": all_dates, "series": {}}
    for name, r in result.items():
        ratio_map = r["ratio"]
        level_map = r.get("level", {})
        shares_map = r.get("total_shares", {})
        payload["series"][name] = {
            "ratio": [round(ratio_map.get(d, None), 3) if d in ratio_map else None for d in all_dates],
            "level": [round(level_map.get(d, None), 1) if d in level_map else None for d in all_dates],
            "shares": [round(shares_map.get(d, None), 0) if d in shares_map else None for d in all_dates],
            "top_codes": r["top_codes"],
            "mcap_now": r["mcap_now"],
        }

    data_json = json.dumps(payload, ensure_ascii=False)
    update_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    return """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>宽基ETF市值 / 指数流通市值</title>
<script src="echarts.min.js"></script>
<style>
  body { font-family: -apple-system, "Microsoft YaHei", sans-serif; margin: 0; background: #0f1117; color: #e6e8eb; }
  header { padding: 24px 32px 8px; }
  h1 { font-size: 20px; margin: 0; }
  .sub { color: #8b93a1; font-size: 13px; margin-top: 6px; }
  .card { margin: 20px 24px 0; background: #14161f; border-radius: 10px; padding: 12px 8px 4px; }
  .card h2 { font-size: 15px; margin: 0 12px 2px; color: #e6e8eb; }
  .latest { color: #8b93a1; font-size: 12px; margin: 0 12px 4px; }
  .chart { width: 100%; height: 340px; }
  .wrap { padding-bottom: 40px; }
</style>
</head>
<body>
<header>
  <h1>宽基ETF市值 占 指数流通市值 比例（前10大ETF之和）</h1>
  <div class="sub">更新时间 __UPDATE_TIME__ · 区间 2023-01 至今 · 流通市值按「当前流通市值 × 指数点位比例」估算 · 数据源：沪深交易所/东财/新浪</div>
</header>
<div class="wrap">
  <div id="charts"></div>
</div>
<script>
const DATA = __DATA__;
const series = DATA.series;
const dates = DATA.dates;
const container = document.getElementById('charts');
const charts = [];

Object.entries(series).forEach(function(entry, idx) {
  var name = entry[0], s = entry[1];
  var latest = '';
  var validIdx = [];
  for (var i = s.ratio.length - 1; i >= 0 && validIdx.length < 2; i--) {
    if (s.ratio[i] != null) validIdx.push(i);
  }
  if (validIdx.length > 0) {
    var i0 = validIdx[0], i1 = validIdx[1];
    var txt = dates[i0] + ' · 占比 ' + s.ratio[i0].toFixed(2) + '%';
    if (i1 != null) {
      var rd = s.ratio[i0] - s.ratio[i1];
      txt += '（' + (rd >= 0 ? '+' : '') + rd.toFixed(2) + 'pp）';
    }
    if (s.shares[i0] != null) {
      txt += '；份额 ' + (s.shares[i0] / 1e8).toFixed(1) + '亿份';
      if (i1 != null && s.shares[i1] != null) {
        var sd = (s.shares[i0] - s.shares[i1]) / 1e8;
        txt += '（' + (sd >= 0 ? '+' : '') + sd.toFixed(1) + '亿份）';
      }
    }
    txt += '；指数点位 ' + s.level[i0].toFixed(1);
    latest = '<div class="latest">' + txt + '</div>';
  }
  var card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = '<h2>' + name + ' · 前10大ETF</h2>' + latest + '<div class="chart" id="chart_' + idx + '"></div>';
  container.appendChild(card);

  var ratioData = s.ratio.map(function(v, i) { return [dates[i], v]; });
  var levelData = s.level.map(function(v, i) { return [dates[i], v]; });

  var chart = echarts.init(document.getElementById('chart_' + idx));
  chart.setOption({
    backgroundColor: 'transparent',
    color: ['#5470c6', '#fac858'],
    tooltip: { trigger: 'axis' },
    legend: { top: 0, right: 10, textStyle: { color: '#c9cfd8' } },
    grid: { left: 60, right: 60, top: 34, bottom: 28 },
    xAxis: { type: 'time', axisLabel: { color: '#8b93a1' }, axisLine: { lineStyle: { color: '#333' } } },
    yAxis: [
      { type: 'value', name: '比例(%)', position: 'left', scale: true, nameTextStyle: { color: '#5470c6' },
        axisLabel: { color: '#5470c6', formatter: '{value}%' }, splitLine: { lineStyle: { color: '#232734' } } },
      { type: 'value', name: '点位', position: 'right', scale: true, nameTextStyle: { color: '#fac858' },
        axisLabel: { color: '#fac858' }, splitLine: { show: false } }
    ],
    series: [
      { name: '比例%', type: 'line', yAxisIndex: 0, showSymbol: false, smooth: true,
        data: ratioData, lineStyle: { width: 2 } },
      { name: '点位', type: 'line', yAxisIndex: 1, showSymbol: false, smooth: true,
        data: levelData, lineStyle: { width: 1.5 } }
    ]
  });
  charts.push(chart);
});

window.addEventListener('resize', function() { charts.forEach(function(c) { c.resize(); }); });
</script>
</body>
</html>
""".replace("__DATA__", data_json).replace("__UPDATE_TIME__", update_time)


if __name__ == "__main__":
    main()
