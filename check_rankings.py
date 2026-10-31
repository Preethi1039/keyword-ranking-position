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


def bare_host(h):
    h = h.strip().lower()
    return h[4:] if h.startswith("www.") else h


def result_host(o):
    """Real website of a result. Google sometimes hides the URL behind a
    google.com/goto?url=... redirect; the site is then only visible in the
    'displayedLink' breadcrumb (e.g. 'https://www.abhasa.in › ocd › ...')."""
    host = bare_host(urlparse(o.get("link", "")).netloc)
    if host.endswith("google.com") and "/goto" in o.get("link", ""):
        shown = o.get("displayedLink", "").split("›")[0].strip()
        if shown:
            return bare_host(urlparse(shown if "://" in shown else "https://" + shown).netloc)
    return host


def host_matches(host, domain):
    return host == domain or host.endswith("." + domain)


def norm_url(u):
    p = urlparse(u.strip().lower())
    host = p.netloc[4:] if p.netloc.startswith("www.") else p.netloc
    return host + p.path.rstrip("/")


def fetch_page(keyword, page, cfg, api_key):
    payload = {"q": keyword, "gl": cfg["country_code"], "hl": cfg["language"],
               "num": 10, "page": page}
    if cfg.get("location"):
        payload["location"] = cfg["location"]
    for attempt in range(3):
        r = requests.post(API_URL, json=payload, timeout=60,
                          headers={"X-API-KEY": api_key, "Content-Type": "application/json"})
        if r.status_code == 200:
            break
        time.sleep(5 * (attempt + 1))
    r.raise_for_status()
    return [o for o in r.json().get("organic", []) if o.get("link")]


def find_ranking(keyword, domain, cfg, api_key):
    """Walk Google result pages (10 per page, as Google serves them) until the
    domain is found or max_position is reached.

    Returns (position, result_dict_or_None, results_checked, raw_log).
    Position = (page - 1) * 10 + place on that page, i.e. what a person sees.
    """
    log, checked = [], 0
    for page in range(1, cfg["max_position"] // 10 + 1):
        organic = fetch_page(keyword, page, cfg, api_key)
        log.append({"page": page, "returned": len(organic),
                    "results": [{"pos": o.get("position"), "site": result_host(o),
                                 "link": o["link"], "displayed": o.get("displayedLink", ""),
                                 "title": o.get("title", "")} for o in organic]})
        if not organic:
            break
        checked = (page - 1) * 10 + len(organic)
        for i, o in enumerate(organic, start=1):
            if host_matches(result_host(o), domain):
                return (page - 1) * 10 + i, o, checked, log
        time.sleep(0.3)
    return None, None, checked, log


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

    rows, failed, raw = [], 0, {}
    for i, (cat, kw, target) in enumerate(items, 1):
        try:
            pos, hit, checked, log = find_ranking(kw, domain, cfg, api_key)
            raw[kw] = {"category": cat, "position": pos, "checked": checked, "pages": log}
            url = ""
            if hit:
                hidden = "/goto" in hit["link"] and "google.com" in hit["link"]
                url = (f"{hit.get('displayedLink', '')}  [title: {hit.get('title', '')}]"
                       if hidden else hit["link"])
            shown, status = (pos if pos else f">{checked}"), "OK"
        except Exception as e:  # keep going; record the failure honestly
            pos, hit, url, shown, status, checked = None, None, "", "Not checked", f"Error: {e}", 0
            failed += 1
        right_page = ""
        if hit and target:
            if url.startswith("http") and "[title:" not in url:
                right_page = "Yes" if norm_url(url) == norm_url(target) else "No - other page"
            else:
                right_page = "Can't tell (Google hid URL)"
        old = prev.get((cat, kw))
        change = ""
        if isinstance(pos, int) and isinstance(old, int):
            change = old - pos  # positive = moved up
        elif isinstance(pos, int) and old is not None:
            change = "New"
        elif pos is None and isinstance(old, int) and status == "OK":
            change = "Lost"
        rows.append([today, cat, kw, shown, "" if old is None else old, change,
                     right_page, url, target, status, checked])
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
                          "Correct page?", "Ranking URL (Google)", "Target page", "Status",
                          "Results checked"],
                 [12, 20, 45, 10, 10, 10, 16, 70, 70, 20, 12])
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
    summary.append([f"Positions are counted among normal (organic) Google results only - ads, map box, "
                    f"'People also ask' and AI answers are not counted. '>{maxp}' = not found in the top {maxp}. "
                    "Raw Google results for every keyword are saved in results/raw/<date>.json."])
    summary.append(["Top N columns are cumulative (Top 10 includes Top 3). "
                    "'Wrong page ranking' = Google ranks a different abhasa.in page than the target page."])

    wb.save(out)
    raw_dir = os.path.join(os.path.dirname(out), "raw")
    os.makedirs(raw_dir, exist_ok=True)
    with open(os.path.join(raw_dir, f"{today}.json"), "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=1)
    print(f"Done: {len(items)} keywords, {failed} not checked. Saved {out}")


if __name__ == "__main__":
    main()
