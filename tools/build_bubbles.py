#!/usr/bin/env python3
"""Build bubbles.json for the Indicator Monitor "Asset Bubbles" section.

Only published index / benchmark prices are used (no single stocks, no
interpolation). Each series is monthly, cut to its bubble window and rebased to
100 at the first month of that window. Run from the repo root:

    python3 tools/build_bubbles.py
"""
import datetime as dt, json, re
from pathlib import Path
from zoneinfo import ZoneInfo
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "tools" / "bubbles_source"

def ms(d): return dt.date(d.year, d.month, 1)

def load_gold(p):
    ws = openpyxl.load_workbook(p, data_only=True)["Monthly Prices"]
    names = [c.value for c in ws[5]]
    col = [i for i, n in enumerate(names) if n and "gold" in str(n).lower()][0]
    out = []
    for r in ws.iter_rows(min_row=7, values_only=True):
        if not r[0] or r[col] in (None, "…"): continue
        m = re.match(r"(\d{4})M(\d{2})", str(r[0]))
        if m: out.append((dt.date(int(m[1]), int(m[2]), 1), float(r[col])))
    return sorted(out)

def load_econstats(p):
    out = []
    for line in p.read_text(errors="replace").splitlines():
        m = re.match(r"^\s*(\d{4})-(\d{2})\s*,", line)
        if not m: continue
        parts = [x.strip() for x in line.split(",")]
        try: out.append((dt.date(int(m[1]), int(m[2]), 1), float(parts[8])))
        except (IndexError, ValueError): pass
    return sorted(out)

def load_yahoo(p, default_tz):
    res = json.loads(p.read_text())["chart"]["result"][0]
    tz = ZoneInfo(res.get("meta", {}).get("exchangeTimezoneName") or default_tz)
    closes = res["indicators"]["quote"][0]["close"]
    by = {}
    for t, c in zip(res["timestamp"], closes):
        if c is None: continue
        by[ms(dt.datetime.fromtimestamp(t, tz).date())] = float(c)
    return sorted(by.items())

D = dt.date
BUBBLES = [
 dict(key="gold", label="Gold (LBMA price)", color="#F0C44E", rows=load_gold(SRC/"CMO-Historical-Data-Monthly.xlsx"),
      start=D(1977,1,1), end=D(1983,12,1), unit="US$/oz",
      source="World Bank Commodity Price Data (Pink Sheet), gold, monthly average US$/troy oz, LBMA-based"),
 dict(key="japan", label="Nikkei 225", color="#4C8DFF", rows=load_yahoo(SRC/"nikkei_225_yahoo_monthly.json","Asia/Tokyo"),
      start=D(1985,1,1), end=D(1992,12,1), unit="pts",
      source="Nikkei 225 (^N225), month-end close, Yahoo Finance"),
 dict(key="asia", label="Hang Seng", color="#F4A0C8", rows=load_econstats(SRC/"hsi_econstats.csv"),
      start=D(1991,1,1), end=D(1998,12,1), unit="pts",
      source="Hang Seng Index (^HSI), month-end close, EconStats"),
 dict(key="tech", label="Nasdaq Composite", color="#E85D5D", rows=load_econstats(SRC/"ixic_econstats.csv"),
      start=D(1995,1,1), end=D(2003,12,1), unit="pts",
      source="Nasdaq Composite (^IXIC), month-end close, EconStats"),
 dict(key="housing", label="PHLX Housing Sector (HGX)", color="#7EC8E8", rows=load_yahoo(SRC/"yahoo_hgx.json","America/New_York"),
      start=D(2002,7,1), end=D(2009,12,1), unit="pts",
      source="PHLX Housing Sector Index (^HGX), month-end close, Yahoo Finance; history begins Jul 2002"),
 dict(key="china", label="Shanghai Composite", color="#D7DCE4", rows=load_yahoo(SRC/"yahoo_sse.json","Asia/Shanghai"),
      start=D(2005,5,1), end=D(2009,12,1), unit="pts",
      source="Shanghai Composite (000001.SS), month-end close, Yahoo Finance"),
 dict(key="biotech", label="Nasdaq Biotechnology (NBI)", color="#3DCE8D", rows=load_yahoo(SRC/"yahoo_nbi.json","America/New_York"),
      start=D(2011,9,1), end=D(2018,12,1), unit="pts",
      source="Nasdaq Biotechnology Index (^NBI), month-end close, Yahoo Finance"),
 dict(key="disruptors", label="NYSE FANG+", color="#FF5A4A", rows=load_yahoo(SRC/"yahoo_fang.json","America/New_York"),
      start=D(2017,1,1), end=D(2026,11,1), unit="pts",
      source="NYSE FANG+ Index (^NYFANG), month-end close, Yahoo Finance; latest point is month-to-date"),
]

series = []
for b in BUBBLES:
    raw = [(d, v) for d, v in b["rows"] if b["start"] <= d <= b["end"]]
    assert raw, b["key"]
    base = raw[0][1]
    pts = [[d.isoformat(), round(v, 4), round(v / base * 100, 4)] for d, v in raw]
    pk = max(pts, key=lambda p: p[2])
    series.append(dict(key=b["key"], label=b["label"], color=b["color"], unit=b["unit"], source=b["source"],
                       start=pts[0][0], end=pts[-1][0], base=round(base, 4),
                       peakDate=pk[0], peakRaw=pk[1], peakRebased=pk[2], points=pts))
    print(f'{b["label"]:30} {pts[0][0]}..{pts[-1][0]} n={len(pts)} peak {pk[0]} rebased {pk[2]:.1f} raw {pk[1]}')

out = dict(
    title="History of asset bubbles since 1977",
    note="Each line is monthly and rebased to 100 at the first month of its window.",
    footnote="Recreated from public index data; not the original slide's exact series. Not investment advice.",
    series=series)
(ROOT / "bubbles.json").write_text(json.dumps(out, separators=(",", ":")))
