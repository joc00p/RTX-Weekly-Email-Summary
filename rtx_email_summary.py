"""
RTX Weekly Team Punch List - Email Summarizer
Reads emails from the Outlook folder, groups by week, and uses Claude AI
to generate detailed per-email summaries and a combined weekly narrative.
"""

import sys
import win32com.client
from collections import defaultdict
from datetime import datetime, timedelta
import urllib.request
import json


def get_week_label(dt):
    if dt.tzinfo is not None:
        dt = dt.astimezone().replace(tzinfo=None)
    monday = dt - timedelta(days=dt.weekday())
    sunday = monday + timedelta(days=6)
    return f"Week of {monday.strftime('%b %d')} - {sunday.strftime('%b %d, %Y')}"


def fetch_emails():
    outlook = win32com.client.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace('MAPI')

    folder = None
    for account in ns.Folders:
        try:
            if 'joel.coopersmith@accenture.com' in account.Name:
                for sub in account.Folders:
                    try:
                        if 'RTX Weekly Team Punch List' in sub.Name:
                            folder = sub
                            break
                    except Exception:
                        continue
        except Exception:
            continue

    if not folder:
        print("ERROR: Could not find 'RTX Weekly Team Punch List' folder.")
        return None

    items = folder.Items
    items.Sort('[ReceivedTime]', True)

    by_week = defaultdict(list)
    for msg in items:
        try:
            received = msg.ReceivedTime
            if received.tzinfo is not None:
                received = received.astimezone().replace(tzinfo=None)
            week = get_week_label(received)
            by_week[week].append({
                'subject': msg.Subject,
                'from': msg.SenderName,
                'received': received.strftime('%a %b %d %I:%M %p'),
                'body': msg.Body.strip(),
            })
        except Exception:
            continue

    return by_week


OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.2"


def call_ollama(prompt):
    payload = json.dumps({
        "model": OLLAMA_MODEL,
        "prompt": prompt,
        "stream": False,
    }).encode()

    req = urllib.request.Request(
        OLLAMA_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=300) as resp:
        data = json.loads(resp.read())
    return data["response"]


def summarize_week(week_label, emails):
    email_blocks = []
    for i, e in enumerate(emails, 1):
        body_excerpt = e['body'][:3000]
        email_blocks.append(
            f"--- Email {i} ---\n"
            f"From: {e['from']}\n"
            f"Subject: {e['subject']}\n"
            f"Received: {e['received']}\n"
            f"Body:\n{body_excerpt}"
        )

    all_emails_text = "\n\n".join(email_blocks)

    prompt = f"""You are a project manager assistant creating a weekly team status report.

Week: {week_label}

Source data:

{all_emails_text}

Generate a structured weekly status report using ONLY the following format. Do not mention emails, senders, punch lists, or that this came from emails. Write as if it is an official team report. Exclude any activities or updates from Joel Coopersmith — do not include him in the report at all.

---

## WEEKLY STATUS REPORT — {week_label}

### Team Updates
For each team member's work, create a section:

**[Full Name] — [Role or Area if inferable]**
- [bullet: key activity or task completed]
- [bullet: key activity or task completed]
- [bullet: issues, blockers, or risks if any]
- [bullet: pending actions or next steps if any]

(Repeat for each person)

---

### Weekly Summary
- [bullet: overall team progress this week]
- [bullet: key accomplishments]
- [bullet: active risks or issues]
- [bullet: items pending or in progress]

---

Use consistent formatting. Use bullet points only — no paragraphs. Be concise and factual."""

    print(f"\n  [Summarizing {week_label} with local {OLLAMA_MODEL}...]", flush=True)
    return call_ollama(prompt)


def print_ai_summary(by_week, max_weeks=None):
    weeks = list(by_week.keys())
    if max_weeks:
        weeks = weeks[:max_weeks]

    for week in weeks:
        emails = by_week[week]
        print(f"\n{'='*70}")
        print(f"  {week}  ({len(emails)} email{'s' if len(emails) != 1 else ''})")
        print(f"{'='*70}")

        summary = summarize_week(week, emails)
        print(summary)
        print()


if __name__ == '__main__':
    max_weeks = int(sys.argv[1]) if len(sys.argv) > 1 else 2

    print(f"Fetching emails from 'RTX Weekly Team Punch List'...")
    by_week = fetch_emails()
    if by_week:
        total = sum(len(v) for v in by_week.values())
        print(f"Found {total} emails across {len(by_week)} weeks.")
        print(f"Summarizing last {max_weeks} week(s) with Claude AI...\n")
        print_ai_summary(by_week, max_weeks=max_weeks)
