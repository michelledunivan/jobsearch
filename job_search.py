"""Daily job search: VP / Head of Insights & Analytics roles.

Targets: senior analytics-executive roles (VP or Head of Insights,
Analytics, Data) as first choice, GM of a business unit as second,
remote-friendly or Phoenix-based. Salary is flagged against a floor
rather than filtered, since many senior postings omit comp data.
Government employers and retail/hospitality GM noise are dropped.
"""

import os
import smtplib
import sys
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests

ADZUNA_APP_ID = os.environ["ADZUNA_APP_ID"]
ADZUNA_APP_KEY = os.environ["ADZUNA_APP_KEY"]
GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
SALARY_FLOOR = int(os.environ.get("SALARY_FLOOR") or 240000)

API_URL = "https://api.adzuna.com/v1/api/jobs/us/search/1"
MAX_DAYS_OLD = 2  # daily run; 2-day window covers late-indexed postings

# priority 1 = first-choice role profile, 2 = second choice.
# title_only requires all words in the job title, so both "VP" and
# "Vice President" spellings need their own query.
ROLE_QUERIES = [
    {"title": "vp analytics", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "vice president analytics", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "vp insights", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "vice president insights", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "head of analytics", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "head of insights", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "vp data", "role": "VP/Head of Analytics & Insights", "priority": 1},
    {"title": "general manager", "role": "General Manager", "priority": 2},
]

SEARCHES = [
    {**query, "mode": mode}
    for query in ROLE_QUERIES
    for mode in ("remote", "phoenix")
]

# GM searches drown in retail/hospitality listings; none of these are
# the business-unit GM profile being targeted.
EXCLUDE_TITLE = [
    "restaurant", "hotel", "retail", "store", "dealership", "automotive",
    "franchise", "gym", "salon", "spa", "warehouse", "car wash",
    "apartment", "property", "branch",
]

# Government is out of scope regardless of role fit.
EXCLUDE_COMPANY = [
    "city of", "county of", "county", "state of", "department of",
    "u.s. ", "federal", "government",
]


def fetch_jobs(search):
    params = {
        "app_id": ADZUNA_APP_ID,
        "app_key": ADZUNA_APP_KEY,
        "results_per_page": 50,
        "title_only": search["title"],
        "max_days_old": MAX_DAYS_OLD,
        "sort_by": "date",
    }
    if search["mode"] == "remote":
        params["what_and"] = "remote"
    else:
        params["where"] = "Phoenix, AZ"
        params["distance"] = 40
    resp = requests.get(API_URL, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json().get("results", [])


def keep(job):
    title = (job.get("title") or "").lower()
    company = (job.get("company", {}).get("display_name") or "").lower()
    if any(word in title for word in EXCLUDE_TITLE):
        return False
    if any(word in company for word in EXCLUDE_COMPANY):
        return False
    return True


def comp_flag(job):
    salary_max = job.get("salary_max")
    if not salary_max:
        return "no salary data — verify before pursuing"
    if salary_max < SALARY_FLOOR:
        return f"below comp bar (max ${salary_max:,.0f})"
    return f"meets comp bar (max ${salary_max:,.0f})"


def to_row(job, search):
    description = job.get("description") or ""
    return {
        "Priority": search["priority"],
        "Role Type": search["role"],
        "Title": job.get("title"),
        "Company": job.get("company", {}).get("display_name"),
        "Location": job.get("location", {}).get("display_name"),
        "Remote Signal": "yes" if "remote" in (job.get("title", "") + description).lower() else "unclear",
        "Compensation": comp_flag(job),
        "Posted": (job.get("created") or "")[:10],
        "Link": job.get("redirect_url"),
    }


def build_email(df, today):
    if df.empty:
        return f"<p>No new VP/Head of Analytics or GM postings in the last {MAX_DAYS_OLD} days that survived filtering. Nothing to review today.</p>"
    parts = [f"<p><b>{len(df)} new posting(s)</b> — VP/Head of Analytics & Insights first, then GM. Comp flagged against ${SALARY_FLOOR:,} floor.</p>"]
    for _, row in df.iterrows():
        parts.append(
            f'<p><b><a href="{row["Link"]}">{row["Title"]}</a></b> — {row["Company"]}<br>'
            f'{row["Location"]} · remote signal: {row["Remote Signal"]} · {row["Compensation"]} · posted {row["Posted"]}</p>'
        )
    return "\n".join(parts)


def send_email(subject, html):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = GMAIL_ADDRESS
    msg.attach(MIMEText(html, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.send_message(msg)


def main():
    rows = []
    for search in SEARCHES:
        try:
            for job in fetch_jobs(search):
                if keep(job):
                    rows.append(to_row(job, search))
        except requests.RequestException as exc:
            print(f"Search failed ({search['title']}, {search['mode']}): {exc}", file=sys.stderr)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link"]).sort_values(["Priority", "Posted"], ascending=[True, False])

    today = date.today().isoformat()
    os.makedirs("output", exist_ok=True)
    out_path = f"output/jobs_{today}.xlsx"
    (df if not df.empty else pd.DataFrame(columns=["Priority", "Role Type", "Title", "Company", "Location", "Remote Signal", "Compensation", "Posted", "Link"])).to_excel(out_path, index=False)
    print(f"Wrote {len(df)} job(s) to {out_path}")

    count = len(df)
    send_email(f"VP Analytics & Insights Job Search — {count} new posting(s) — {today}", build_email(df, today))
    print("Email sent.")


if __name__ == "__main__":
    main()
