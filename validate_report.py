"""Independent validator for results/rankings.xlsx.

Re-derives every position from the raw Google data (results/raw/<date>.json)
WITHOUT using the tracker's own matching code, then checks strict rules:

 R1  Every keyword in data/keywords.xlsx is in the report (none missing/extra)
 R2  Every keyword was actually checked (Status = OK, raw data exists)
 R3  Position = (page-1)*10 + place of the FIRST abhasa.in result in raw data
 R4  Disguised google.com/goto links are resolved via the displayed address
 R5  ">N" is used only when abhasa.in is absent from ALL raw results, and
     N equals the number of results really checked
 R6  A full check (>= 95 results) was done for every "not found" keyword
 R7  Pages were read in order 1,2,3... with no gaps, max 10 results per page
 R8  Ranking URL in the report matches the raw result at that position
 R9  "Correct page?" agrees with Ranking URL vs Target page
 R10 Previous/Change agree with the previous run in History
"""
import json
import sys
from urllib.parse import urlparse

from openpyxl import load_workbook

DOMAIN = "abhasa.in"


def host_of(url):
    h = urlparse(url if "://" in url else "https://" + url).netloc.lower()
    return h[4:] if h.startswith("www.") else h


def real_site(r):
    link = r.get("link", "")
    if "google.com/goto" in link:
        return host_of(r.get("displayed", "").split("›")[0].strip())
    return host_of(link)


def is_ours(site):
    return site == DOMAIN or site.endswith("." + DOMAIN)


def norm(u):
    p = urlparse(u.strip().lower())
    return host_of(p.netloc) + p.path.rstrip("/")


def split_runs(hist):
    """History rows -> list of runs (a new run starts when a keyword repeats)."""
    runs, cur, seen = [], [], set()
    for h in hist:
        key = (h[1], h[2])
        if key in seen:
            runs.append(cur)
            cur, seen = [], set()
        cur.append(h)
        seen.add(key)
    if cur:
        runs.append(cur)
    return runs


def derive(pages):
    """Independent re-derivation of position from raw pages."""
    true_pos, hit, n_checked, ended = None, None, 0, False
    for p in pages:
        res = p["results"]
        if res:
            n_checked = (p["page"] - 1) * 10 + len(res)
        if len(res) < 10:
            ended = True  # Google had no more results
        for i, x in enumerate(res, 1):
            if is_ours(real_site(x)):
                return (p["page"] - 1) * 10 + i, x, n_checked, ended
    return true_pos, hit, n_checked, ended


def main(run_no=None, raw_file=None):
    import glob
    sys.path.insert(0, ".")
    import check_rankings as c  # only to read the keyword list
    expected = {(cat, kw) for cat, kw, _ in c.load_keywords("data/keywords.xlsx")}
    wb = load_workbook("results/rankings.xlsx")
    runs = split_runs(list(wb["History"].iter_rows(min_row=2, values_only=True)))
    idx = (int(run_no) - 1) if run_no else len(runs) - 1
    run, prev_run = runs[idx], (runs[idx - 1] if idx > 0 else None)
    date = run[0][0]
    if not raw_file:
        cands = sorted(glob.glob(f"results/raw/{date}*.json"))
        raw_file = cands[-1] if cands else None
    raw = json.load(open(raw_file, encoding="utf-8")) if raw_file else {}
    is_latest = idx == len(runs) - 1
    latest = {(r[1], r[2]): r for r in wb["Latest"].iter_rows(min_row=2, values_only=True)} if is_latest else {}
    prev = {(h[1], h[2]): h[3] for h in (prev_run or []) if h[3] != "Not checked"}

    fails, checks = [], 0

    def check(rule, ok, msg):
        nonlocal checks
        checks += 1
        if not ok:
            fails.append(f"{rule}: {msg}")

    got = {(h[1], h[2]) for h in run}
    check("R1", got == expected, f"missing={sorted(expected - got)} extra={sorted(got - expected)}")

    for h in run:
        d, cat, kw, pos, url, status = h[:6]
        tag = f"[{kw}]"
        check("R2", status == "OK" and kw in raw, f"{tag} status={status}, raw={'yes' if kw in raw else 'no'}")
        if kw not in raw or status != "OK":
            continue
        pages = raw[kw]["pages"]
        check("R7", [p["page"] for p in pages] == list(range(1, len(pages) + 1))
              and all(p["returned"] <= 10 for p in pages), f"{tag} page order/size wrong")
        true_pos, hit, n_checked, ended = derive(pages)
        if true_pos:
            check("R3", pos == true_pos, f"{tag} report={pos} raw={true_pos}")
            if "google.com/goto" in hit["link"]:
                check("R4", hit["displayed"].split("›")[0].strip() in (url or ""), f"{tag} hidden link not resolved: {url}")
            else:
                check("R8", url == hit["link"], f"{tag} url {url} != raw {hit['link']}")
        else:
            check("R5", pos == f">{n_checked}", f"{tag} report={pos} but raw has no abhasa.in in {n_checked}")
            check("R6", n_checked >= 95 or ended, f"{tag} stopped after {n_checked} results without reaching the end")

        if is_latest:
            L = latest.get((cat, kw))
            check("R1", L is not None, f"{tag} missing from Latest")
            if not L:
                continue
            check("R3", L[3] == pos, f"{tag} Latest={L[3]} History={pos}")
            target, right = L[8], L[6]
            if target and url and url.startswith("http") and "[title:" not in url:
                want = "Yes" if norm(url) == norm(target) else "No - other page"
                check("R9", right == want, f"{tag} correct-page={right}, expected {want}")
            old = prev.get((cat, kw))
            check("R10", (L[4] if L[4] is not None else "") == ("" if old is None else old),
                  f"{tag} previous={L[4]}, previous run says {old}")
            if isinstance(pos, int) and isinstance(old, int):
                ch = str(L[5]).replace("+", "")
                check("R10", ch.lstrip("-").isdigit() and int(ch) == old - pos,
                      f"{tag} change={L[5]}, expected {old - pos}")

    print(f"Run validated: #{idx + 1} of {len(runs)}, date {date}, raw data: {raw_file}")
    print(f"Checks performed: {checks}")
    if fails:
        print(f"FAILED: {len(fails)}")
        for f in fails:
            print("  -", f)
    else:
        print("ALL CHECKS PASSED")
    return not fails


if __name__ == "__main__":
    sys.exit(0 if main(*sys.argv[1:]) else 1)
