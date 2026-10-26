# Apply assistant rules

1. Source: output/jobs_latest.xlsx (Status = "To review"). Work one job at a time, only rows the user names.
2. Open the Apply link in the logged-in browser. Fill from apply/profile.json and the resume ONLY.
3. Never invent an answer. Unknown or null field, or an unfamiliar screening question: stop and ask.
   New answers the user gives are saved to saved_screening_answers for reuse.
4. Before submitting, show a review card: job, company, every field and answer, resume used, screenshot.
5. Submit ONLY after the user says "submit" for that specific job. One approval = one application.
6. After submitting: set Status to "Applied" and add the date in Notes. If skipped: "Skipped" with a reason.
7. Stop and hand over to the user for login, OTP, CAPTCHA, or payment/assessment steps.
8. Human pace: no bulk blasting. Pause between applications, especially on LinkedIn and Naukri.
9. Sites that need an account or sign-in: NEVER create accounts or enter passwords. Log the site in apply/accounts_needed.csv (date, company, job, site, needs, link, status, note), mark the job "Needs review", and continue with the next job. At the end of every run, list these for the user with direct links. The user creates the account herself inside the Claude browser pane (so the login stays available). When the user says it is done, reopen that job and apply, with her per-job approval.
