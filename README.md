# Job Finder

Collects backend / full stack / agentic-AI openings (0-3 yrs or experience not stated), removes
duplicates, and writes an Excel file with Apply links. Run it whenever you like.

## Setup (once)

```bash
cd job-finder
pip3 install requests beautifulsoup4 openpyxl
```

## Run any time

```bash
python3 jobs.py
```

Output goes to `output/jobs_<date>_<time>.xlsx` (and a `.json` copy for the later resume/cover-letter step).
New jobs since your last run are marked **NEW**. The `Status` column has a dropdown (To review, Applied...).
The **Search Links** sheet has ready-made searches for LinkedIn, Naukri, Wellfound, HiringCafe and YC.

Options: `python3 jobs.py --sources himalayas,career_pages --days 14`

## Daily schedule (optional, Mac)

```bash
chmod +x schedule_mac.sh
./schedule_mac.sh install 09:00     # remove with: ./schedule_mac.sh remove
```

## LinkedIn and Naukri (via your email alerts)

LinkedIn and Naukri restrict scraping, so this reads their job-alert emails instead.
1. Create job alerts on both sites (daily email).
2. Make a Gmail app password (Google Account > Security > 2-Step Verification > App passwords).
3. In Terminal: `export JOBS_EMAIL_USER=you@gmail.com JOBS_EMAIL_APP_PASSWORD=xxxxxxxxxxxxxxxx`
4. Set `"email_alerts": true` under `sources` in `config.json`.

For the scheduled run, add those two variables to the plist (`EnvironmentVariables`) or keep this source manual.

## Customising

Edit `config.json`: title keywords, excluded words (senior, lead...), locations, max years, and the
company lists under `career_pages` (use the slug from the company's Greenhouse/Lever/Ashby URL,
e.g. `boards.greenhouse.io/<slug>`). Wrong slugs are skipped with a message.

## Known limits

- YC and HiringCafe have no official API; those fetchers are best effort and may need updating if
  the sites change. Wellfound blocks bots, so use its link in the Search Links sheet.
- Experience is read from the text. If a listing doesn't state it, it's kept and shown as "Not stated".
- Applying: this tool only finds and lists jobs. The tailor-and-apply step comes next and will
  always stop for your review before anything is submitted.
