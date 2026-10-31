"""Weekly Google rank checker.

Reads keywords (and their target page URLs) from EVERY sheet of an Excel file,
fetches the LIVE Google results for each keyword (via the Serper.dev API),
finds where your domain ranks, and saves results/rankings.xlsx with:
  - Summary : position buckets per category for this run
  - Latest  : this run, with previous position and change
  - History : every run, appended

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
from openpyxl.styles import Alignment, Font, PatternFill

API_URL = "https://google.serper.dev/search"


def load_config():
    with open("config.json", encoding="utf-8") as f:
        return json.load(f)


def is_url(v):
    return isinstance(v, str) and v.strip().lower().startswith(("http://", "https://"))


def load_keywords(path):
    """Return [(category, keyword, target_url)] from every sheet.

    In each row the cell that is a URL is the target page and the first
    non-URL text cell is the keyword, whatever the column order. The header
    row (no URL in it) is skipped automatically.
    """
    wb = load_workbook(path, read_only=True)
    items, seen = [], set()
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cells = [c for c in row if c is not None and str(c).strip()]
            urls = [str(c).strip() for c in cells if is_url(c)]
            texts = [str(c).strip() for c in cells if not is_url(c)]
            if not texts:
                continue
            if not urls and any(h in texts[0].lower() for h in ("keyword", "page name", "url")):
                continue  # header row
            kw = texts[0]
            key = (ws.title, kw.lower())
            if key in seen:
                continue
            seen.add(key)
            items.append((ws.title, kw, urls[0] if urls else ""))
    return items


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


def norm_url(u):
    p = urlparse(u.strip().lower())
    host = p.netloc[4:] if p.netloc.startswith("www.") else p.netloc
    return host + p.path.rstrip("/")


def fetch_results(keyword, cfg, api_key):
    """Return the ordered list of organic result URLs (live Google SERP)."""
    links, page = [], 1
    while len(links) < cfg["max_position"]:
        payload = {"q": keyword, "gl": cfg["country_code"], "hl": cfg["language"],
                   "num": 100, "page": page}
        if cfg.get("location"):
            payload["location"] = cfg["location"]
        for attempt in range(3):
            r = requests.post(API_URL, json=payload, timeout=60,
                              headers={"X-API-KEY": api_key, "Content-Type": "application/json"})
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


def previous_positions(path):
    """Latest recorded position per (category, keyword), from History."""
    prev = {}
    if not os.path.exists(path):
        return prev
    wb = load_workbook(path)
    if "History" not in wb.sheetnames:
        return prev
    for row in wb["History"].iter_rows(min_row=2, values_only=True):
        if row[3] != "Not checked":
            prev[(row[1], row[2])] = row[3]  # chronological: last wins
    return prev


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
GREEN, RED = Font(color="008000", bold=True), Font(color="C00000", bold=True)


def write_header(ws, headers, widths):
    ws.append(headers)
    for c, w in zip(ws[1], widths):
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def bucket(pos):
    if not isinstance(pos, int):
        return None
    for limit in (3, 10, 20, 50, 100):
        if pos <= limit:
            return limit
    return None


def main():
    api_key = os.environ.get("SERPER_API_KEY")
    if not api_key:
        sys.exit("SERPER_API_KEY is not set.")
    cfg = load_config()
    domain = clean_domain(cfg["domain"])
    items = load_keywords(cfg["keywords_file"])
    out = cfg["output_file"]
    os.makedirs(os.path.dirname(out), exist_ok=True)
    today = date.today().isoformat()
    prev = previous_positions(out)
    maxp = cfg["max_position"]

    rows, failed = [], 0
    for i, (cat, kw, target) in enumerate(items, 1):
        try:
            links = fetch_results(kw, cfg, api_key)
            pos, url = find_position(links, domain)
            shown, status = (pos if pos else f">{maxp}"), "OK"
        except Exception as e:  # keep going; record the failure honestly
            pos, url, shown, status = None, "", "Not checked", f"Error: {e}"
            failed += 1
        right_page = ""
        if url and target:
            right_page = "Yes" if norm_url(url) == norm_url(target) else "No - other page"
        old = prev.get((cat, kw))
        change = ""
        if isinstance(pos, int) and isinstance(old, int):
            change = old - pos  # positive = moved up
        elif isinstance(pos, int) and old is not None:
            change = "New"
        elif pos is None and isinstance(old, int) and status == "OK":
            change = "Lost"
        rows.append([today, cat, kw, shown, "" if old is None else old, change,
                     right_page, url, target, status])
        print(f"[{i}/{len(items)}] {cat} | {kw}: {shown}")
        time.sleep(0.5)

    if os.path.exists(out):
        wb = load_workbook(out)
        for name in ("Summary", "Latest"):
            if name in wb.sheetnames:
                del wb[name]
        summary = wb.create_sheet("Summary", 0)
        latest = wb.create_sheet("Latest", 1)
        hist = wb["History"]
    else:
        wb = Workbook()
        summary = wb.active
        summary.title = "Summary"
        latest = wb.create_sheet("Latest")
        hist = wb.create_sheet("History")
        write_header(hist, ["Date", "Category", "Keyword", "Position", "Ranking URL", "Status"],
                     [12, 20, 45, 10, 70, 20])

    write_header(latest, ["Date", "Category", "Keyword", "Position", "Previous", "Change",
                          "Correct page?", "Ranking URL (Google)", "Target page", "Status"],
                 [12, 20, 45, 10, 10, 10, 16, 70, 70, 20])
    for r in rows:
        latest.append(r)
        n = latest.max_row
        ch = latest.cell(row=n, column=6)
        if isinstance(r[5], int) and r[5] > 0:
            ch.value, ch.font = f"+{r[5]}", GREEN
        elif (isinstance(r[5], int) and r[5] < 0) or r[5] == "Lost":
            ch.font = RED
        if r[6].startswith("No"):
            latest.cell(row=n, column=7).font = RED
        hist.append([r[0], r[1], r[2], r[3], r[7], r[9]])
    latest.auto_filter.ref = latest.dimensions

    # Summary: counts per category
    heads = ["Category", "Keywords", "Top 3", "Top 10", "Top 20", "Top 50", "Top 100",
             f"Not in top {maxp}", "Not checked", "Wrong page ranking", "Moved up", "Moved down"]
    write_header(summary, heads, [22] + [12] * (len(heads) - 1))
    summary.insert_rows(1)
    summary["A1"] = f"Google rankings for {domain} - {cfg.get('location') or cfg['country_code'].upper()} - checked {today}"
    summary["A1"].font = Font(bold=True, size=13)
    summary.freeze_panes = "A3"
    cats = list(dict.fromkeys(r[1] for r in rows)) + ["ALL"]
    for cat in cats:
        rs = [r for r in rows if cat == "ALL" or r[1] == cat]
        ints = [r[3] for r in rs if isinstance(r[3], int)]
        line = [cat, len(rs)] + [sum(1 for p in ints if p <= lim) for lim in (3, 10, 20, 50, 100)]
        line += [sum(1 for r in rs if r[9] == "OK" and not isinstance(r[3], int)),
                 sum(1 for r in rs if r[3] == "Not checked"),
                 sum(1 for r in rs if r[6].startswith("No")),
                 sum(1 for r in rs if isinstance(r[5], int) and r[5] > 0),
                 sum(1 for r in rs if (isinstance(r[5], int) and r[5] < 0) or r[5] == "Lost")]
        summary.append(line)
        if cat == "ALL":
            for c in summary[summary.max_row]:
                c.font = Font(bold=True)
    summary.append([])
    summary.append(["Top N columns are cumulative (Top 10 includes Top 3). "
                    "'Wrong page ranking' = Google ranks a different abhasa.in page than the target page."])

    wb.save(out)
    print(f"Done: {len(items)} keywords, {failed} not checked. Saved {out}")


if __name__ == "__main__":
    main()
