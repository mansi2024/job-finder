#!/usr/bin/env python3
"""
Job finder: pulls openings from several sources, keeps only roles that match your
filters (backend / full stack / agentic AI, 0-3 years or unspecified), removes
duplicates, and writes one Excel sheet with Apply links.

Run any time:   python3 jobs.py
Options:        python3 jobs.py --sources himalayas,career_pages --days 14
"""
import argparse
import datetime as dt
import email
import hashlib
import imaplib
import json
import os
import re
import sys
import urllib.parse
from email.header import decode_header
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
UA = {"User-Agent": "Mozilla/5.0 (personal job-search script)"}
TIMEOUT = 20


# ----------------------------------------------------------------- helpers
def log(msg):
    print(msg, flush=True)


def get_json(url, **kw):
    r = requests.get(url, headers=UA, timeout=TIMEOUT, **kw)
    r.raise_for_status()
    return r.json()


def strip_html(s):
    return BeautifulSoup(s or "", "html.parser").get_text(" ", strip=True)


def make_job(title, company, location, url, source, posted="", text="", remote=None):
    return {
        "title": (title or "").strip(),
        "company": (company or "").strip(),
        "location": (location or "").strip(),
        "url": (url or "").strip(),
        "source": source,
        "posted": posted or "",
        "text": (text or "")[:30000],
        "remote": remote,
    }


# ----------------------------------------------------------------- sources
def fetch_himalayas(cfg):
    jobs = []
    for q in cfg["himalayas_queries"]:
        try:
            data = get_json("https://himalayas.app/jobs/api/search",
                            params={"q": q, "sort": "recent"})
        except Exception as e:
            log(f"  himalayas '{q}': {e}")
            continue
        for j in data.get("jobs", []):
            locs = j.get("locationRestrictions") or []
            # Remote jobs limited to specific countries are useless unless India is one of them.
            if locs and not any("india" in l.lower() for l in locs):
                continue
            loc = ", ".join(locs) if locs else "Worldwide (remote)"
            pub = j.get("pubDate")
            if isinstance(pub, (int, float)):  # API gives a Unix timestamp
                pub = dt.datetime.fromtimestamp(pub).strftime("%Y-%m-%d")
            jobs.append(make_job(
                j.get("title"), j.get("companyName"), loc,
                j.get("applicationLink") or j.get("guid"), "Himalayas",
                str(pub or "")[:10], strip_html(j.get("description", "")),
                remote=True))
    return jobs




def fetch_career_pages(cfg):
    jobs = []
    cp = cfg["career_pages"]
    for slug in cp.get("greenhouse", []):
        try:
            # content=true includes each job's description, needed for the experience filter
            d = get_json(f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs",
                         params={"content": "true"})
        except Exception as e:
            log(f"  greenhouse/{slug}: {e}")
            continue
        for j in d.get("jobs", []):
            import html as _html
            body = strip_html(_html.unescape(j.get("content", "")))
            jobs.append(make_job(j.get("title"), slug.title(),
                                 (j.get("location") or {}).get("name", ""),
                                 j.get("absolute_url"), f"Career page ({slug})",
                                 str(j.get("updated_at", ""))[:10], body))
    for slug in cp.get("lever", []):
        try:
            d = get_json(f"https://api.lever.co/v0/postings/{slug}", params={"mode": "json"})
        except Exception as e:
            log(f"  lever/{slug}: {e}")
            continue
        for j in d:
            ts = j.get("createdAt")
            posted = dt.datetime.fromtimestamp(ts / 1000).strftime("%Y-%m-%d") if ts else ""
            jobs.append(make_job(j.get("text"), slug.title(),
                                 (j.get("categories") or {}).get("location", ""),
                                 j.get("hostedUrl"), f"Career page ({slug})", posted,
                                 j.get("descriptionPlain", "")))
    for slug in cp.get("ashby", []):
        try:
            d = get_json(f"https://api.ashbyhq.com/posting-api/job-board/{slug}")
        except Exception as e:
            log(f"  ashby/{slug}: {e}")
            continue
        for j in d.get("jobs", []):
            jobs.append(make_job(j.get("title"), slug.title(), j.get("location", ""),
                                 j.get("jobUrl"), f"Career page ({slug})",
                                 str(j.get("publishedAt", ""))[:10],
                                 j.get("descriptionPlain", ""),
                                 remote=j.get("isRemote")))
    return jobs


def fetch_email_alerts(cfg):
    """Reads LinkedIn / Naukri job-alert emails from your inbox (read-only).
    Set up job alerts on those sites first. Needs an app password (not your real one)."""
    ec = cfg["email_alerts"]
    user, pwd = os.environ.get(ec["user_env"]), os.environ.get(ec["password_env"])
    if not user or not pwd:
        log(f"  email alerts: set {ec['user_env']} and {ec['password_env']} to enable")
        return []
    jobs = []
    since = (dt.date.today() - dt.timedelta(days=ec["lookback_days"])).strftime("%d-%b-%Y")
    try:
        M = imaplib.IMAP4_SSL(ec["imap_host"])
        M.login(user, pwd)
        M.select("INBOX", readonly=True)
        for sender, src in [("linkedin.com", "LinkedIn"), ("naukri.com", "Naukri")]:
            _, ids = M.search(None, f'(SINCE {since} FROM "{sender}")')
            for num in ids[0].split():
                _, msg_data = M.fetch(num, "(RFC822)")
                msg = email.message_from_bytes(msg_data[0][1])
                for part in msg.walk():
                    if part.get_content_type() != "text/html":
                        continue
                    html = part.get_payload(decode=True).decode(errors="ignore")
                    for a in BeautifulSoup(html, "html.parser").find_all("a", href=True):
                        h, t = a["href"], a.get_text(" ", strip=True)
                        if len(t) < 8:
                            continue
                        if ("/jobs/view/" in h and src == "LinkedIn") or \
                           ("job-listings" in h and src == "Naukri"):
                            jobs.append(make_job(t, "", "", h, f"{src} alert"))
        M.logout()
    except Exception as e:
        log(f"  email alerts: {e}")
    return jobs


def fetch_linkedin(cfg):
    """LinkedIn's public (no-login) job listing endpoints, at a low rate. Pulls each posting's
    description too so the experience filter can see '5+ years' style requirements."""
    import time
    lc = cfg.get("linkedin", {})
    secs = max(cfg["max_age_days"], 1) * 86400
    found = {}
    for q in cfg["queries"]:
        for start in range(0, lc.get("pages_per_query", 4) * 10, 10):
            try:
                r = requests.get(
                    "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search",
                    params={"keywords": q, "location": "India", "f_TPR": f"r{secs}",
                            "sortBy": "DD", "start": start},
                    headers=UA, timeout=TIMEOUT)
                if r.status_code == 429:
                    log("  linkedin: rate limited, stopping early")
                    break
                r.raise_for_status()
            except Exception as e:
                log(f"  linkedin '{q}' start={start}: {e}")
                break
            cards = BeautifulSoup(r.text, "html.parser").find_all("li")
            n = 0
            for li in cards:
                a = li.select_one("a.base-card__full-link")
                if not a:
                    continue
                n += 1
                g = lambda s: (li.select_one(s).get_text(" ", strip=True) if li.select_one(s) else "")
                t = li.find("time")
                url = a["href"].split("?")[0]
                found.setdefault(url, dict(
                    title=g(".base-search-card__title"), company=g(".base-search-card__subtitle"),
                    location=g(".job-search-card__location"),
                    posted=(t.get("datetime") if t else "") or "", url=url))
            time.sleep(1.5)
            if n == 0:
                break
    inc = [k.lower() for k in cfg["title_include"]]
    exc = [k.lower() for k in cfg["title_exclude"]]
    cand = [r for r in found.values()
            if any(k in r["title"].lower() for k in inc) and not any(k in r["title"].lower() for k in exc)
            and not title_blocked(r["title"], cfg)][:lc.get("max_descriptions", 250)]
    log(f"  linkedin: {len(found)} listings, reading {len(cand)} descriptions (slow on purpose)")
    jobs = []
    for r in cand:
        text = ""
        m = re.search(r"(\d{6,})/?$", r["url"])
        if m:
            try:
                d = requests.get(f"https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{m.group(1)}",
                                 headers=UA, timeout=TIMEOUT)
                if d.status_code == 429:
                    log("  linkedin: rate limited while reading descriptions")
                elif d.ok:
                    el = BeautifulSoup(d.text, "html.parser").select_one(".show-more-less-html__markup")
                    text = el.get_text(" ", strip=True) if el else ""
            except Exception:
                pass
            time.sleep(1.0)
        jobs.append(make_job(r["title"], r["company"], r["location"], r["url"], "LinkedIn",
                             r["posted"][:10], text))
    return jobs


SOURCES = {  # plain web/API sources; the browser sources live in browser_sources.py
    "himalayas": fetch_himalayas,
    "career_pages": fetch_career_pages,
    "email_alerts": fetch_email_alerts,
    "linkedin": fetch_linkedin,
}
BROWSER_SOURCE_NAMES = ["naukri", "hiringcafe", "yc"]


# ----------------------------------------------------------------- filtering
YEARS_RANGE = re.compile(r"(\d{1,2})\s*(?:\+)?\s*(?:-|–|to)\s*(\d{1,2})\s*\+?\s*(?:years|yrs|year)", re.I)
YEARS_MIN = re.compile(r"(?:minimum|at least|min\.?)?\s*(\d{1,2})\s*\+\s*(?:years|yrs|year)|(\d{1,2})\s*(?:years|yrs)\s*(?:of)?\s*(?:experience|exp)", re.I)


YEARS_ANY = re.compile(
    r"(\d{1,2})\s*(?:\+|plus)?\s*(?:(?:-|–|—|to)\s*(\d{1,2})\s*\+?)?\s*(?:or more\s*)?(?:years?|yrs?)\b", re.I)


def experience_range(text):
    """Experience the posting asks for, as (min_years, max_years), or None if not stated.

    Looks at every 'N years' / 'N-M years' / 'N+ years' mention that sits near the word
    'experience' and keeps the highest minimum, so one '8+ years' line is enough to
    exclude a posting even if another line says '2 years'."""
    best = None
    for m in YEARS_ANY.finditer(text):
        if text[m.end(): m.end() + 8].lower().lstrip().startswith("ago"):
            continue
        near = text[max(0, m.start() - 80): m.end() + 80].lower()
        if "experience" not in near and "exp" not in near:
            continue
        lo = int(m.group(1))
        hi = int(m.group(2)) if m.group(2) else None
        if hi is not None and hi < lo:  # e.g. "2020-22 years" noise
            continue
        if lo > 15:  # '20 years of company history' style noise
            continue
        if best is None or lo > best[0]:
            best = (lo, hi)
    return best


def category_of(job):
    t = (job["title"] + " " + job["text"][:800]).lower()
    if re.search(r"agentic|\bllm\b|genai|generative ai|ai engineer|applied ai|machine learning", t):
        return "Agentic / AI"
    if re.search(r"full[\s-]?stack", t):
        return "Full Stack"
    return "Backend"


def location_ok(job, cfg):
    loc = (job["location"] + " " + job["title"]).lower()
    if any(c in loc for c in cfg["locations_india"]):
        return True
    if cfg["include_global_remote"] and (job["remote"] or "remote" in loc or "worldwide" in loc
                                         or "anywhere" in loc):
        # "Remote - USA", "Remote, Canada" etc. are not open to you; plain "Remote" is.
        if any(re.search(r"\b" + re.escape(w) + r"\b", loc) for w in cfg["non_india_regions"]):
            return False
        return True
    return job["location"] == ""  # unknown location: keep, let the user judge


def too_old(job, cfg):
    p = job["posted"][:10]
    try:
        d = dt.date.fromisoformat(p)
    except ValueError:
        return False
    return (dt.date.today() - d).days > cfg["max_age_days"]


def title_blocked(title, cfg):
    """Roles you can't do (e.g. iOS / Swift)."""
    return any(re.search(p, title, re.I) for p in cfg.get("title_exclude_regex", []))


def keep(job, cfg):
    title = job["title"].lower()
    if title_blocked(job["title"], cfg):
        return False
    if "mobile" in title and re.search(r"\bios\b|swiftui|\bswift\b", job["text"][:3000], re.I):
        return False
    if not job["title"] or not job["url"]:
        return False
    if not any(k in title for k in cfg["title_include"]):
        return False
    if any(k in title for k in cfg["title_exclude"]):
        return False
    if not location_ok(job, cfg) or too_old(job, cfg):
        return False
    exp = experience_range(job["title"] + " " + job["text"])
    job["experience"] = "Not stated"
    if exp:
        lo, hi = exp
        job["experience"] = f"{lo}-{hi} yrs" if hi else f"{lo}+ yrs"
        if lo > cfg["max_years_experience"]:
            return False
    return True


def days_ago(posted):
    """Whole days since the posting date, or None if the date is unknown."""
    try:
        return max((dt.date.today() - dt.date.fromisoformat(posted[:10])).days, 0)
    except (ValueError, TypeError):
        return None


def company_tier(company, cfg):
    """'Top-tier' / 'Well-known' if the company is in the curated lists in config.json."""
    c = re.sub(r"[^a-z0-9 ]", "", company.lower()).strip()
    if not c:
        return ""
    for label, key in (("Top-tier", "top_tier"), ("Well-known", "well_known")):
        for name in cfg["renowned_companies"][key]:
            n = re.sub(r"[^a-z0-9 ]", "", name.lower()).strip()
            if c == n or c.startswith(n + " "):
                return label
    return ""


def dedupe(jobs):
    seen, out = set(), []
    for j in jobs:
        # Drop tracking params only; some sites (Greenhouse) identify the job in the query string.
        url_key = re.sub(r"[?&](utm_[^&#]*|ref=[^&#]*)", "", j["url"]).split("#")[0].rstrip("/?").lower()
        k = hashlib.md5((url_key or f"{j['company']}{j['title']}").encode()).hexdigest()
        k2 = hashlib.md5(f"{j['company']}|{j['title']}|{j['location']}".lower().encode()).hexdigest()
        if k in seen or (j["company"] and k2 in seen):
            continue
        seen.add(k)
        seen.add(k2)
        j["id"] = k
        out.append(j)
    return out


# ----------------------------------------------------------------- output
def search_links():
    q = lambda s: urllib.parse.quote_plus(s)
    rows = []
    for kw in ["backend engineer", "java developer", "full stack developer", "ai engineer agentic"]:
        rows += [
            ("LinkedIn", kw, f"https://www.linkedin.com/jobs/search/?keywords={q(kw)}&location=India&f_TPR=r864000"),
            ("Naukri", kw, f"https://www.naukri.com/{kw.replace(' ', '-')}-jobs?experience=3"),
            ("Wellfound", kw, "https://wellfound.com/jobs"),
            ("HiringCafe", kw, f"https://hiring.cafe/?searchState={q(json.dumps({'searchQuery': kw}))}"),
            ("YC Work at a Startup", kw, "https://www.workatastartup.com/jobs"),
        ]
    return rows


def load_previous_status(path):
    """Apply-link -> (status, notes) from the last sheet, so your Applied/Skip marks survive re-runs."""
    prev = {}
    if not Path(path).exists():
        return prev
    try:
        from openpyxl import load_workbook
        wbk = load_workbook(path)
        for name in ("Jobs", "Experience Not Stated"):
            if name not in wbk.sheetnames:
                continue
            ws = wbk[name]
            head = [c.value for c in ws[1]]
            a, st, nt = head.index("Apply"), head.index("Status"), head.index("Notes")
            for row in ws.iter_rows(min_row=2):
                link = row[a].hyperlink.target if row[a].hyperlink else None
                if link:
                    prev[link] = (row[st].value or "To review", row[nt].value or "")
    except Exception as e:  # a damaged old sheet must never block a new run
        log(f"  (could not read previous statuses: {e})")
    return prev


def skill_score(job, cfg):
    blob = (job["title"] + " " + job["text"]).lower()
    return sum(1 for k in cfg.get("my_skills", []) if re.search(r"\b" + re.escape(k.lower()) + r"\b", blob))


def write_excel(jobs, path, new_ids, prev=None):
    prev = prev or {}
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    head = ["New?", "Days ago", "Posted", "Category", "Title", "Company", "Company tier",
            "Location", "Experience", "Skill match", "Source", "Apply", "Check pay", "Status", "Notes"]
    widths = [7, 10, 12, 14, 44, 22, 14, 26, 12, 11, 22, 9, 11, 14, 30]
    tier_rank = {"Top-tier": 0, "Well-known": 1, "": 2}
    # Latest first (unknown dates last); within the same day, better companies first.
    ordered = sorted(jobs, key=lambda x: (x["days_ago"] is None, x["days_ago"] or 0,
                                          tier_rank[x["tier"]], x["company"]))

    def fill(sheet, rows):
        sheet.append(head)
        for c in sheet[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        for j in rows:
            sheet.append(["NEW" if j["id"] in new_ids else "", j["days_ago"], j["posted"],
                          j["category"], j["title"], j["company"], j["tier"], j["location"],
                          j["experience"], j["skills"], j["source"], "Apply",
                          "Pay info" if j["company"] else "", *prev.get(j["url"], ("To review", ""))])
            r = sheet.max_row
            sheet.cell(row=r, column=12).hyperlink = j["url"]
            sheet.cell(row=r, column=12).font = Font(color="0563C1", underline="single")
            if j["company"]:
                q = urllib.parse.quote_plus(j["company"])
                sheet.cell(row=r, column=13).hyperlink = f"https://www.ambitionbox.com/search?q={q}"
                sheet.cell(row=r, column=13).font = Font(color="0563C1", underline="single")
            if j["tier"] == "Top-tier":
                sheet.cell(row=r, column=7).fill = PatternFill("solid", fgColor="C6EFCE")
            elif j["tier"] == "Well-known":
                sheet.cell(row=r, column=7).fill = PatternFill("solid", fgColor="FFEB9C")
        dv = DataValidation(type="list",
                            formula1='"To review,Shortlisted,Applied,Interview,Rejected,Skip"')
        sheet.add_data_validation(dv)
        dv.add(f"N2:N{max(sheet.max_row, 2)}")
        for col, w in zip("ABCDEFGHIJKLMNO", widths):
            sheet.column_dimensions[col].width = w
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

    stated = [j for j in ordered if j["experience"] != "Not stated"]
    fill(ws, stated)
    top = wb.create_sheet("Top Companies")
    fill(top, [j for j in stated if j["tier"]])
    # Postings that don't mention experience: kept apart so you can judge them separately.
    fill(wb.create_sheet("Experience Not Stated"), [j for j in ordered if j["experience"] == "Not stated"])

    ws2 = wb.create_sheet("Search Links")
    ws2.append(["Portal", "Search", "Open"])
    for c in ws2[1]:
        c.font = Font(bold=True)
    for portal, kw, url in search_links():
        ws2.append([portal, kw, "Open"])
        ws2.cell(row=ws2.max_row, column=3).hyperlink = url
        ws2.cell(row=ws2.max_row, column=3).font = Font(color="0563C1", underline="single")
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 26
    wb.save(path)


# ----------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=str(HERE / "config.json"))
    ap.add_argument("--sources", help="comma list, e.g. himalayas,career_pages (default: from config)")
    ap.add_argument("--days", type=int, help="override max_age_days")
    ap.add_argument("--login", action="store_true",
                    help="open Naukri / YC / HiringCafe / LinkedIn once so you can log in (session is saved)")
    ap.add_argument("--import-file", action="append",
                    help="JSON file of extra listings (title, company, location, url, posted); repeatable")
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text())
    if args.login:
        from browser_sources import login_session
        return login_session(cfg)
    if args.days:
        cfg["max_age_days"] = args.days
    chosen = args.sources.split(",") if args.sources else [k for k, v in cfg["sources"].items() if v]
    api_chosen = [c for c in chosen if c in SOURCES]            # "--sources none" skips live fetching
    browser_chosen = [c for c in chosen if c in BROWSER_SOURCE_NAMES]

    all_jobs = []
    for f in args.import_file or []:  # listings collected another way (e.g. through a browser)
        for r in json.loads(Path(f).read_text()):
            all_jobs.append(make_job(r["title"], r["company"], r.get("location", ""), r["url"],
                                     r.get("source", "Imported"), r.get("posted", ""),
                                     r.get("text", ""), r.get("remote")))
    for name in api_chosen:
        log(f"Fetching {name} ...")
        try:
            got = SOURCES[name](cfg)
        except Exception as e:  # one broken source must not stop the run
            log(f"  {name} failed: {e}")
            got = []
        log(f"  {len(got)} raw listings")
        all_jobs += got

    if browser_chosen:
        from browser_sources import run_browser_sources
        for r in run_browser_sources(browser_chosen, cfg):
            all_jobs.append(make_job(r["title"], r["company"], r.get("location", ""), r["url"],
                                     r["source"], r.get("posted", ""), r.get("text", ""), r.get("remote")))

    matched = dedupe([j for j in all_jobs if keep(j, cfg)])
    for j in matched:
        j["category"] = category_of(j)
        j["days_ago"] = days_ago(j["posted"])
        j["tier"] = company_tier(j["company"], cfg)
        j["skills"] = skill_score(j, cfg)

    out_dir = HERE / cfg["output_dir"]
    out_dir.mkdir(exist_ok=True)
    seen_file = out_dir / "seen.json"
    seen = set(json.loads(seen_file.read_text())) if seen_file.exists() else set()
    new_ids = {j["id"] for j in matched} - seen
    seen_file.write_text(json.dumps(sorted(seen | {j["id"] for j in matched})))

    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
    xlsx = out_dir / f"jobs_{stamp}.xlsx"
    latest = out_dir / "jobs_latest.xlsx"
    write_excel(matched, xlsx, new_ids, load_previous_status(latest))
    import shutil
    shutil.copy(xlsx, latest)  # the file to keep open; your Status/Notes carry into the next run
    # JSON copy: input for the later "tailor resume + cover letter and apply" step
    (out_dir / f"jobs_{stamp}.json").write_text(json.dumps(
        [{k: v for k, v in j.items() if k != "text"} | {"description": j["text"]} for j in matched],
        indent=1))
    log(f"\n{len(matched)} matching jobs ({len(new_ids)} new) -> {xlsx}")


if __name__ == "__main__":
    sys.exit(main())
