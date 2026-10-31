# Weekly Keyword Rank Tracker

Checks the live Google position of every keyword in `data/keywords.xlsx` for your
domain, every Monday, and saves/emails `results/rankings.xlsx`.

## One-time setup (≈10 minutes)

1. **Serper API key** – sign up at https://serper.dev and copy your API key.
2. **Create a GitHub repo** (private is fine) and upload all files from this folder.
3. **Put your keywords** in `data/keywords.xlsx` – one keyword per row in column A
   (a "Keyword" header in row 1 is optional).
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
- **Latest** – this week: Keyword, Domain, Position, Previous, Change (+ = moved up), Ranking URL
- **History** – every week's results, so you can track trends

`>100` means not in the top 100. `Not checked` means the API call failed for that
keyword – nothing is ever estimated.

To change the day/time, edit the `cron` line in `.github/workflows/weekly-rank-check.yml`.
