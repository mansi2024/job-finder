# Job Finder

Finds fresh job postings for **your** role, experience level and locations, removes duplicates, and
writes them to an Excel sheet with Apply links. An optional second half helps you apply: it picks the
best matches, fills forms from **your** profile and resume, and always stops for your approval before
anything is submitted.

Everything personal (your resume, profile, results, application log) stays on your machine and is
git-ignored, so you can fork or clone this repo without leaking anything.

## What you get

| Sheet | What it holds |
|---|---|
| `Jobs` | Postings that match your role, with a stated experience you qualify for, newest first |
| `Top Companies` | The same, limited to companies in your `renowned_companies` list |
| `Experience Not Stated` | Dated postings that don't mention years of experience |
| `No Posted Date` | Postings with no posted date (can't be age-filtered, so judge them yourself) |
| `Search Links` | Ready-made searches for LinkedIn, Naukri, Wellfound, HiringCafe, YC |

Each row has a `Status` dropdown (To review, Shortlisted, Applied, ...). Your marks survive re-runs.
Jobs new since your last run are marked **NEW**.

## 1. Install (once)

You need Python 3.9+.

```bash
git clone <this repo's URL>
cd job-finder
pip3 install -r requirements.txt
python3 -m playwright install chromium     # only needed for LinkedIn / Naukri / HiringCafe
```

## 2. Make it yours: `config.json`

Open `config.json` and change these. The defaults target a 0-2 year Java/backend/AI engineer in India.

| Key | What to set |
|---|---|
| `max_years_experience` | Highest *minimum* experience a posting may ask for. `2` keeps "0-2" and "2-3 yrs", drops "3+" |
| `max_age_days` | Only postings newer than this many days. Also see `browser.hiringcafe_days`, `browser.naukri_job_age` (Naukri accepts 1, 3, 7, 15, 30) |
| `title_include` | Words a job title must contain. **Replace with your target roles**, e.g. `["data engineer", "etl", "spark"]` or `["frontend", "react"]` |
| `title_exclude` | Titles to skip (senior, lead, intern, ...). Remove `senior` etc. if you are senior |
| `title_exclude_regex` | Regex patterns for stacks you can't do (the default skips iOS/Swift) |
| `queries` | Search phrases sent to job boards. Match them to your roles |
| `himalayas_queries` | Search phrases for Himalayas (remote jobs) |
| `locations_india` | Cities you will work from. Any other country works too: just list its cities, and edit `non_india_regions` so your own region isn't blocked |
| `include_global_remote` | `true` to also keep remote jobs with no country restriction |
| `my_skills` | Your skills. Used for the "Skill match" score and ranking |
| `renowned_companies` | Companies you want highlighted as Top-tier / Well-known |
| `career_pages` | Companies to read directly from Greenhouse / Lever / Ashby. Use the slug in the company's job-board URL, e.g. `boards.greenhouse.io/<slug>`. Wrong slugs are skipped with a message |
| `sources` | Switch sources on/off: `himalayas`, `career_pages`, `naukri`, `linkedin`, `hiringcafe`, `yc`, `email_alerts` |

### Example: a fresher frontend developer in Bengaluru

```json
"max_years_experience": 1,
"title_include": ["frontend", "front-end", "react", "ui engineer"],
"queries": ["react developer", "frontend engineer"],
"himalayas_queries": ["frontend engineer", "react developer"],
"locations_india": ["bengaluru", "bangalore"],
"my_skills": ["react", "javascript", "typescript", "css", "html", "redux"]
```

Also edit `category_of()` in `jobs.py` (search for `"Agentic / AI"`) if you want your own category labels.

## 3. Find jobs

```bash
python3 jobs.py
```

Output: `output/jobs_latest.xlsx` (plus dated copies and a `.json`). Open it in Excel / Numbers / Sheets.

Options:

```bash
python3 jobs.py --sources himalayas,career_pages     # only some sources
python3 jobs.py --days 5                             # override max_age_days
```

LinkedIn, Naukri and HiringCafe use a real browser window (Playwright). The first time, log in to the
sites in the window that opens; the login is kept in `~/.jobfinder-profile` for next runs.

### Optional: use job-alert emails instead of scraping

1. Create daily job alerts on LinkedIn / Naukri.
2. Make a Gmail app password (Google Account > Security > 2-Step Verification > App passwords).
3. `export JOBS_EMAIL_USER=you@gmail.com JOBS_EMAIL_APP_PASSWORD=xxxxxxxxxxxxxxxx`
4. Set `"email_alerts": true` under `sources`.

### Optional: run it every day (macOS)

```bash
chmod +x schedule_mac.sh
./schedule_mac.sh install 09:00      # remove with: ./schedule_mac.sh remove
```

On Linux use cron (`0 9 * * * cd /path/to/job-finder && python3 jobs.py`); on Windows use Task Scheduler.

## 4. Apply (optional, human-approved)

The `apply/` folder holds the apply workflow. It is designed to be driven by an AI coding assistant
(Codex, Claude Code, ...) that can read files and control a browser, following `apply/APPLY_RULES.md`.

Set up once:

1. `cp apply/profile.example.json apply/profile.json` and fill in **your** details. Anything left
   `null` is never guessed: the assistant stops and asks you.
2. Put your resume PDF in `apply/` and point `resume_pdf` in `profile.json` to it. Keep your
   editable source (e.g. LaTeX) in `apply/master/` if you want per-job tailored resumes.
3. Edit `apply/apply_config.json`:
   - `categories`: which sheet categories to apply to
   - `experience_filter`: your range (`accept_if_min_years_at_most`, `accept_if_max_years_at_least`)
   - `title_exclude`, `location_allow`, `location_exclude`, `company_exclude`: your preferences
   - `max_per_run`: how many jobs per session

Each session:

```bash
python3 apply/queue.py --limit 5             # shows the best unapplied matches as JSON
python3 apply/mark.py --url <job-url> --status Applied --note "via company site"
```

`queue.py` ranks by company tier, category and skill match, and skips anything already in
`apply/applied_log.csv`. `mark.py` updates the sheet (with a backup in `output/backup/`) and the log.

The rules the assistant must follow (in `apply/APPLY_RULES.md`): never invent an answer, show a review
card before each submission, submit only after you say "submit" for that job, never create accounts or
type passwords, and stop for OTP / CAPTCHA. Review every application yourself.

## Learn the code with the Codex app (free to download)

Want to understand or change how this works? Download the **Codex app** and open this folder in it, then
ask things like "explain how `jobs.py` filters by experience" or "add a new job source".

- Download: <https://developers.openai.com/codex/app> (macOS and Windows)
- Sign in with a ChatGPT account or an OpenAI API key. Usage limits depend on your plan.

Any other coding assistant (Claude Code, Cursor, VS Code + Copilot) works too.

## Project layout

```
jobs.py              main script: fetch, filter, dedupe, write Excel
browser_sources.py   LinkedIn / Naukri / HiringCafe via Playwright
config.json          your role, locations, sources
schedule_mac.sh      optional daily run on macOS
apply/               optional apply workflow (queue.py, mark.py, profile, rules)
output/              results (git-ignored)
```

## Known limits

- Scraped sites change their pages; fetchers for YC, HiringCafe, LinkedIn and Naukri are best effort and
  may need updates. Wellfound blocks bots, so use its link in `Search Links`.
- Experience is read from posting text. If it isn't stated the job is kept and shown as "Not stated".
- Respect each site's terms of use and keep a human pace when applying; don't bulk-submit.

## Privacy

`.gitignore` keeps your profile, resume, application log and results out of git. If you committed
personal files earlier, they are still in git history: remove them with `git filter-repo` before making
a repo public.
