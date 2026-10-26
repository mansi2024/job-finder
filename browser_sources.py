"""
Browser-based sources: Naukri, HiringCafe, Y Combinator (Work at a Startup) and LinkedIn.

These sites need a real browser session, so this uses Playwright with a dedicated profile
folder (default ~/.jobfinder-profile). Log in ONCE with:   python3 jobs.py --login
After that the saved session is reused on every run.

It reads search-result pages the way you would (a few pages per search, with pauses). It
never clicks Apply and never submits anything.

One-time setup:  pip3 install playwright && python3 -m playwright install chromium
"""
import datetime as dt
import os
import re
import time
import urllib.parse

try:
    from playwright.sync_api import sync_playwright
except ImportError:  # handled by the caller
    sync_playwright = None


def log(msg):
    print(msg, flush=True)


# --------------------------------------------------------------- date helpers
def age_to_date(text):
    """'Just now', '5h', '2d', '1w', '3 Days Ago', '30+ days ago' -> 'YYYY-MM-DD' (or '')."""
    t = (text or "").strip().lower()
    today = dt.date.today()
    if not t:
        return ""
    if "just now" in t or "few hours" in t or "today" in t or re.fullmatch(r"\d+\s*(h|m|min)", t):
        return today.isoformat()
    m = re.search(r"(\d+)\s*\+?\s*(mo|month|w|week|d|day)", t)
    if not m:
        return ""
    n, unit = int(m.group(1)), m.group(2)
    days = n * (30 if unit.startswith("mo") or unit == "month" else 7 if unit.startswith("w") else 1)
    return (today - dt.timedelta(days=days)).isoformat()


# --------------------------------------------------------------- JS extractors
# (Selectors checked against the live pages on 2026-10-25. If a site redesigns, fix them here.)
JS_HIRINGCAFE = r"""
() => {
  const seen = new Set(), out = [];
  for (const a of document.querySelectorAll('a[href^="/job/"]')) {
    const h = a.getAttribute('href'); if (seen.has(h)) continue; seen.add(h);
    const card = a.parentElement.parentElement.parentElement;
    const L = card.innerText.split("\n").map(s => s.trim()).filter(Boolean);
    const yoe = (card.innerText.match(/(\d{1,2})\s*(?:\+|-\s*\d{1,2})?\s*YOE/i) || [])[0] || "";
    const co = (L.find(s => /^[^:]{2,60}:\s/.test(s)) || "").split(":")[0];
    out.push({age: L[0], title: L[1], location: L[2], workplace: L[3], company: co, yoe, href: h});
  }
  return out;
}"""

JS_NAUKRI = r"""
() => {
  const seen = new Set(), out = [];
  for (const c of document.querySelectorAll('.srp-jobtuple-wrapper')) {
    const t = c.querySelector('a.title'); if (!t || seen.has(t.href)) continue; seen.add(t.href);
    const g = s => ((c.querySelector(s) || {}).innerText || "").trim();
    out.push({title: t.innerText.trim(), href: t.href, company: g('a.comp-name'),
              exp: g('.expwdth'), location: g('.locWdth'), posted: g('.job-post-day'),
              salary: g('.sal-wrap'), skills: [...c.querySelectorAll('.tags-gt li')].map(x => x.innerText).join(', ')});
  }
  return out;
}"""

JS_YC = r"""
() => {
  const seen = new Set(), out = [];
  for (const a of document.querySelectorAll('a[href*="/jobs/"]')) {
    const h = a.getAttribute('href'); if (!/\/jobs\/\d+/.test(h) || seen.has(h)) continue; seen.add(h);
    let el = a; for (let i = 0; i < 4 && el.parentElement && el.innerText.length < 60; i++) el = el.parentElement;
    const L = el.innerText.split("\n").map(s => s.trim()).filter(Boolean);
    out.push({href: h, company: L[0], tagline: L[1], title: L[2], meta: L[3]});
  }
  return out;
}"""

JS_LINKEDIN_SEARCH = r"""
async ([q, starts]) => {
  const out = [];
  for (const start of starts) {
    try {
      const r = await fetch(`/jobs-guest/jobs/api/seeMoreJobPostings/search?keywords=${encodeURIComponent(q)}&location=India&f_E=2%2C3&f_TPR=r1209600&sortBy=DD&start=${start}`);
      if (!r.ok) continue;
      const doc = new DOMParser().parseFromString(await r.text(), "text/html");
      doc.querySelectorAll("li").forEach(li => {
        const a = li.querySelector("a.base-card__full-link"); if (!a) return;
        const g = s => ((li.querySelector(s) || {}).textContent || "").trim();
        out.push({title: g(".base-search-card__title"), company: g(".base-search-card__subtitle"),
                  location: g(".job-search-card__location"),
                  posted: (li.querySelector("time") || {}).getAttribute?.("datetime") || "",
                  url: a.href.split("?")[0].replace("https://in.linkedin.com", "https://www.linkedin.com")});
      });
    } catch (e) {}
    await new Promise(r => setTimeout(r, 1500));
  }
  return out;
}"""

JS_LINKEDIN_DESC = r"""
async (urls) => {
  const res = {};
  for (const u of urls) {
    try {
      const r = await fetch(u); if (!r.ok) { res[u] = null; if (r.status === 429) break; continue; }
      const doc = new DOMParser().parseFromString(await r.text(), "text/html");
      res[u] = ((doc.querySelector(".show-more-less-html__markup") || {}).textContent || "").replace(/\s+/g, " ");
    } catch (e) { res[u] = null; }
    await new Promise(r => setTimeout(r, 1200));
  }
  return res;
}"""


# --------------------------------------------------------------- browser plumbing
def _context(p, bcfg):
    profile = os.path.expanduser(bcfg.get("profile_dir", "~/.jobfinder-profile"))
    return p.chromium.launch_persistent_context(
        profile, headless=bcfg.get("headless", False),
        viewport={"width": 1280, "height": 900})


def login_session(cfg):
    """Open the sites so you can log in once. Press Enter in the terminal when done."""
    if sync_playwright is None:
        raise SystemExit("Run: pip3 install playwright && python3 -m playwright install chromium")
    with sync_playwright() as p:
        ctx = _context(p, {**cfg["browser"], "headless": False})
        for url in ["https://www.naukri.com/nlogin/login", "https://www.workatastartup.com/",
                    "https://hiringcafe.com/", "https://www.linkedin.com/login"]:
            page = ctx.new_page()
            page.goto(url)
        input("\nLog in to the sites in the browser window (you type your own passwords), "
              "then press Enter here to save the session... ")
        ctx.close()
    log("Session saved.")


def _scroll(page, times=6, pause=1.2):
    for _ in range(times):
        page.mouse.wheel(0, 3000)
        time.sleep(pause)


# --------------------------------------------------------------- sources
def fetch_hiringcafe(page, cfg):
    jobs = []
    days = cfg["browser"].get("hiringcafe_days", 7)
    for q in cfg["queries"]:
        state = urllib.parse.quote(
            '{"searchQuery":%s,"dateFetchedPastNDays":%d,"sortBy":"date"}' % (_jstr(q), days))
        try:
            page.goto(f"https://hiringcafe.com/?searchState={state}", wait_until="domcontentloaded")
            page.wait_for_selector('a[href^="/job/"]', timeout=20000)
            _scroll(page, cfg["browser"].get("scrolls", 6))
            rows = page.evaluate(JS_HIRINGCAFE)
        except Exception as e:
            log(f"  hiringcafe '{q}': {e}")
            continue
        for r in rows:
            exp = f" Experience: {r['yoe'].replace('YOE', 'years')} of experience." if r["yoe"] else ""
            jobs.append(dict(
                title=r["title"], company=re.sub(r"(NYSE|NASDAQ|NSE|LSE|BSE|/|\s)+$", "", r["company"]),
                location=r["location"], url="https://hiringcafe.com" + r["href"], source="HiringCafe",
                posted=age_to_date(r["age"]), text=f"{r['workplace']}.{exp}",
                remote=("remote" in (r["workplace"] or "").lower()) or None))
        time.sleep(2)
    return jobs


def fetch_naukri(page, cfg):
    jobs = []
    b = cfg["browser"]
    yrs = b.get("naukri_experience", 2)
    for q in cfg["queries"]:
        slug = re.sub(r"[^a-z0-9]+", "-", q.lower()).strip("-")
        for pg in range(1, b.get("naukri_pages", 3) + 1):
            suffix = f"-{pg}" if pg > 1 else ""
            url = (f"https://www.naukri.com/{slug}-jobs{suffix}?k={urllib.parse.quote(q)}"
                   f"&experience={yrs}&jobAge={b.get('naukri_job_age', 7)}&sort=f")
            try:
                page.goto(url, wait_until="domcontentloaded")
                page.wait_for_selector(".srp-jobtuple-wrapper", timeout=20000)
                rows = page.evaluate(JS_NAUKRI)
            except Exception as e:
                log(f"  naukri '{q}' p{pg}: {e}")
                break
            for r in rows:
                exp = r["exp"] if re.search(r"yr", r["exp"], re.I) else ""
                exp_txt = f"Experience: {exp.replace('Yrs', 'years')} of experience." if exp else ""
                jobs.append(dict(
                    title=r["title"], company=r["company"], location=r["location"], url=r["href"],
                    source="Naukri", posted=age_to_date(r["posted"]),
                    text=f"{exp_txt} {r['skills']} {r['salary']}"))
            time.sleep(3)
    return jobs


def fetch_yc(page, cfg):
    jobs = []
    for role in ["software-engineer", "backend-engineer", "full-stack-engineer", "machine-learning-engineer"]:
        try:
            page.goto(f"https://www.workatastartup.com/jobs/l/{role}", wait_until="domcontentloaded")
            page.wait_for_selector('a[href*="/jobs/"]', timeout=20000)
            _scroll(page, 4)
            rows = page.evaluate(JS_YC)
        except Exception as e:
            log(f"  yc {role}: {e}")
            continue
        for r in rows:
            meta = r.get("meta") or ""
            loc = meta.split("·")[1].strip() if "·" in meta else ""
            jobs.append(dict(title=r["title"], company=r["company"], location=loc,
                             url="https://www.workatastartup.com" + r["href"], source="Y Combinator",
                             posted="", text=r.get("tagline", "")))
        time.sleep(2)
    return jobs


def fetch_linkedin(page, cfg):
    """Uses LinkedIn's public (no-login) job listing pages at a low rate, then reads each
    posting's description so the experience filter can see '5+ years' style requirements."""
    page.goto("https://www.linkedin.com/jobs", wait_until="domcontentloaded")
    found = {}
    for q in cfg["queries"]:
        try:
            for r in page.evaluate(JS_LINKEDIN_SEARCH, [q, [0, 25]]):
                found.setdefault(r["url"], r)
        except Exception as e:
            log(f"  linkedin '{q}': {e}")
    # Cheap title pre-filter so we only open postings worth reading.
    inc = [k.lower() for k in cfg["title_include"]]
    exc = [k.lower() for k in cfg["title_exclude"]]
    keep = [r for r in found.values()
            if any(k in r["title"].lower() for k in inc) and not any(k in r["title"].lower() for k in exc)]
    log(f"  linkedin: {len(found)} listings, reading {len(keep)} descriptions (slow on purpose)")
    desc = {}
    urls = [r["url"] for r in keep]
    for i in range(0, len(urls), 10):
        try:
            desc.update(page.evaluate(JS_LINKEDIN_DESC, urls[i:i + 10]) or {})
        except Exception as e:
            log(f"  linkedin descriptions: {e}")
    return [dict(title=r["title"], company=r["company"], location=r["location"], url=r["url"],
                 source="LinkedIn", posted=r["posted"][:10], text=desc.get(r["url"]) or "")
            for r in keep]


def _jstr(s):
    import json
    return json.dumps(s)


BROWSER_SOURCES = {
    "naukri": fetch_naukri,
    "hiringcafe": fetch_hiringcafe,
    "yc": fetch_yc,
}


def run_browser_sources(names, cfg):
    """Returns a list of job dicts (title, company, location, url, source, posted, text)."""
    if sync_playwright is None:
        log("  Playwright not installed: pip3 install playwright && python3 -m playwright install chromium")
        return []
    jobs = []
    with sync_playwright() as p:
        ctx = _context(p, cfg["browser"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        for name in names:
            log(f"Fetching {name} (browser) ...")
            try:
                got = BROWSER_SOURCES[name](page, cfg)
            except Exception as e:
                log(f"  {name} failed: {e}")
                got = []
            log(f"  {len(got)} raw listings")
            jobs += got
        ctx.close()
    return jobs
