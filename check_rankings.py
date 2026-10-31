"""Weekly Google rank checker.

Reads keywords from an Excel file, fetches the LIVE Google results for each
keyword at the configured location (via the Serper.dev API), finds where your
domain ranks, and appends the results to results/rankings.xlsx.

Nothing is estimated: if a keyword cannot be checked, the row says so.
"""
import json
import os
import sys
import time
from datetime import date
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

API_URL = "https://google.serper.dev/search"


def load_config():
    with open("config.json", encoding="utf-8") as f:
        return json.load(f)


def load_keywords(path):
    wb = load_workbook(path, read_only=True)
    ws = wb.active
    keywords, seen = [], set()
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if not row or row[0] is None:
            continue
        kw = str(row[0]).strip()
        if i == 0 and kw.lower() in ("keyword", "keywords", "query", "search term"):
            continue  # header row
        if kw and kw.lower() not in seen:
            seen.add(kw.lower())
            keywords.append(kw)
    return keywords


def clean_domain(d):
    d = d.strip().lower()
    if "://" in d:
        d = urlparse(d).netloc
    d = d.split("/")[0]
    return d[4:] if d.startswith("www.") else d


def host_matches(link, domain):
    host = urlparse(link).netloc.lower()
    host = host[4:] if host.startswith("www.") else host
    return host == domain or host.endswith("." + domain)


def fetch_results(keyword, cfg, api_key):
    """Return the ordered list of organic result URLs (live Google SERP)."""
    links, page = [], 1
    while len(links) < cfg["max_position"]:
        payload = {
            "q": keyword,
            "gl": cfg["country_code"],
            "hl": cfg["language"],
            "location": cfg["location"],
            "num": 100,
            "page": page,
        }
        for attempt in range(3):
            r = requests.post(API_URL, json=payload, timeout=60,
                              headers={"X-API-KEY": api_key,
                                       "Content-Type": "application/json"})
            if r.status_code == 200:
                break
            time.sleep(5 * (attempt + 1))
        r.raise_for_status()
        organic = r.json().get("organic", [])
        if not organic:
            break
        links.extend(o["link"] for o in organic if o.get("link"))
        if len(organic) < 10:
            break
        page += 1
    return links[: cfg["max_position"]]


def find_position(links, domain):
    for pos, link in enumerate(links, start=1):
        if host_matches(link, domain):
            return pos, link
    return None, ""


def previous_positions(path, domain):
    """Latest recorded position per keyword for this domain, from History."""
    prev = {}
    if not os.path.exists(path):
        return prev
    wb = load_workbook(path)
    if "History" not in wb.sheetnames:
        return prev
    for row in wb["History"].iter_rows(min_row=2, values_only=True):
        d, kw, dom, pos = row[0], row[1], row[2], row[3]
        if dom == domain:
            prev[kw] = pos  # rows are chronological, so last one wins
    return prev


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def write_header(ws, headers, widths):
    ws.append(headers)
    for c, w in zip(ws[1], widths):
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def main():
    api_key = os.environ.get("SERPER_API_KEY")
    if not api_key:
        sys.exit("SERPER_API_KEY is not set.")
    cfg = load_config()
    domains = [clean_domain(cfg["domain"])] + [clean_domain(c) for c in cfg.get("competitors", [])]
    keywords = load_keywords(cfg["keywords_file"])
    out = cfg["output_file"]
    os.makedirs(os.path.dirname(out), exist_ok=True)
    today = date.today().isoformat()
    prev = {d: previous_positions(out, d) for d in domains}

    rows, failed = [], 0
    for i, kw in enumerate(keywords, 1):
        try:
            links = fetch_results(kw, cfg, api_key)
            status = "OK"
        except Exception as e:  # keep going; record the failure honestly
            links, status = None, f"Not checked: {e}"
            failed += 1
        for d in domains:
            if links is None:
                pos, url = None, ""
                shown = "Not checked"
            else:
                pos, url = find_position(links, d)
                shown = pos if pos else f">{cfg['max_position']}"
            old = prev[d].get(kw)
            change = ""
            if isinstance(pos, int) and isinstance(old, int):
                change = old - pos  # positive = moved up
            elif isinstance(pos, int) and old is not None:
                change = "New"
            rows.append([today, kw, d, shown, old if old is not None else "", change, url, status])
        print(f"[{i}/{len(keywords)}] {kw}: {rows[-len(domains)][3]}")
        time.sleep(0.5)

    # Workbook: "Latest" (this run) + "History" (every run, appended)
    if os.path.exists(out):
        wb = load_workbook(out)
        if "Latest" in wb.sheetnames:
            del wb["Latest"]
        latest = wb.create_sheet("Latest", 0)
        hist = wb["History"]
    else:
        wb = Workbook()
        latest = wb.active
        latest.title = "Latest"
        hist = wb.create_sheet("History")
        write_header(hist, ["Date", "Keyword", "Domain", "Position", "Ranking URL", "Status"],
                     [12, 40, 25, 10, 60, 25])

    write_header(latest, ["Date", "Keyword", "Domain", "Position", "Previous", "Change", "Ranking URL", "Status"],
                 [12, 40, 25, 10, 10, 10, 60, 25])
    green, red = Font(color="008000", bold=True), Font(color="C00000", bold=True)
    for r in rows:
        latest.append(r)
        ch = latest.cell(row=latest.max_row, column=6)
        if isinstance(r[5], int) and r[5] > 0:
            ch.value, ch.font = f"+{r[5]}", green
        elif isinstance(r[5], int) and r[5] < 0:
            ch.font = red
        hist.append([r[0], r[1], r[2], r[3], r[6], r[7]])
    wb.save(out)
    print(f"Done: {len(keywords)} keywords, {failed} not checked. Saved {out}")


if __name__ == "__main__":
    main()
