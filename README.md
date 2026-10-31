# Weekly Keyword Rank Tracker

Checks the live Google position of every keyword in `data/keywords.xlsx` for your
domain, every Monday, and saves/emails `results/rankings.xlsx`.

## One-time setup (≈10 minutes)

1. **Serper API key** – sign up at https://serper.dev and copy your API key.
2. **Create a GitHub repo** (private is fine) and upload all files from this folder.
3. **Keywords** live in `data/keywords.xlsx`. Every sheet is read; each sheet is a
   category. Each row needs a keyword and (optionally) its target page URL, in any
   column order. To add/remove keywords, replace this file.
4. **Edit `config.json`** – your `domain`, optional `competitors`, `location`
   (e.g. `"Chennai, Tamil Nadu, India"`), `country_code` (`in`) and `language` (`en`).
5. **Add secrets** – repo → Settings → Secrets and variables → Actions → New secret:
   - `SERPER_API_KEY` – your Serper key (required)
   - `MAIL_USERNAME` – Gmail address to send from (optional, for email)
   - `MAIL_PASSWORD` – a Gmail **App Password** (Google Account → Security → App passwords)
   - `MAIL_TO` – where to send the report
6. **Test it** – Actions tab → "Weekly rank check" → **Run workflow**.

After that it runs automatically every Monday morning (IST).

## Output (`results/rankings.xlsx`)
- **Summary** – per category: how many keywords are in Top 3 / 10 / 20 / 50 / 100, moved up/down, wrong page ranking
- **Latest** – this week: Category, Keyword, Position, Previous, Change (+ = moved up),
  Correct page? (does Google rank the target page or a different abhasa.in page), Ranking URL, Target page
- **History** – every week's results, so you can track trends

`>100` means not in the top 100. `Not checked` means the API call failed for that
keyword – nothing is ever estimated.

To change the day/time, edit the `cron` line in `.github/workflows/weekly-rank-check.yml`.

## How positions are counted
- Google is read page by page, 10 results per page, up to `max_position` (100).
  Position = (page - 1) x 10 + place on the page - the same number a person counts.
- Only normal (organic) results count: ads, the map box, "People also ask" and
  AI Overviews are not counted.
- Google sometimes hides a result's address behind a `google.com/goto` link; the
  tracker then identifies the site from the address Google displays under the title.
  For those rows "Correct page?" says "Can't tell" and the title is shown instead.
- Searching stops as soon as your domain is found, so well-ranked keywords cost
  1-2 searches; keywords not in the top 100 cost 10.
- Raw Google results behind every number: `results/raw/<date>.json`.
