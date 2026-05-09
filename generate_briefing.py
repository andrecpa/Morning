"""
Generates the daily morning briefing HTML by pulling live calendar and email
data from Microsoft Graph, then asking Claude to write the email.
"""
import os
import json
import datetime
import zoneinfo
import requests
import anthropic

EASTERN = zoneinfo.ZoneInfo("America/New_York")


def get_access_token() -> str:
    url = (
        f"https://login.microsoftonline.com/{os.environ['MS_TENANT_ID']}"
        "/oauth2/v2.0/token"
    )
    resp = requests.post(url, data={
        "grant_type": "client_credentials",
        "client_id": os.environ["MS_CLIENT_ID"],
        "client_secret": os.environ["MS_CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default",
    })
    resp.raise_for_status()
    return resp.json()["access_token"]


def get_emails(token: str, user_email: str, hours_back: int = 24) -> list:
    since = (
        datetime.datetime.now(datetime.timezone.utc)
        - datetime.timedelta(hours=hours_back)
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{user_email}/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={
            "$filter": f"receivedDateTime ge {since}",
            "$select": "subject,from,receivedDateTime,bodyPreview,importance,isRead",
            "$orderby": "receivedDateTime desc",
            "$top": 30,
        },
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


def get_calendar_events(token: str, user_email: str, days_ahead: int = 5) -> list:
    now_et = datetime.datetime.now(EASTERN)
    start = now_et.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + datetime.timedelta(days=days_ahead + 1)

    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{user_email}/calendarView",
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="Eastern Standard Time"',
        },
        params={
            "startDateTime": start.strftime("%Y-%m-%dT%H:%M:%S"),
            "endDateTime": end.strftime("%Y-%m-%dT%H:%M:%S"),
            "$select": "subject,start,end,location,attendees,organizer,bodyPreview,showAs",
            "$orderby": "start/dateTime",
            "$top": 50,
        },
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


STYLE_GUIDE = """
HTML email structure (max-width 620px, white card on #f0f2f5 background):

1. HEADER: background #1a1a2e, white h1 "Morning Briefing", muted subtitle with date,
   small all-caps company name above.

2. EMAIL SUMMARY section:
   - Each item: border-left 4px, border-radius 0 6px 6px 0, colored background tint.
   - Urgency tiers:
       URGENT     → border #d9534f, bg #fdf7f7  (action required today / security alerts)
       ATTENTION  → border #e8a020, bg #fdfaf4  (needs review or reply soon)
       INFO       → border #3d7ab5, bg #f4f7fd  (useful context, no urgent action)
       LOW        → border #888888, bg #f9f9f9  (FYI only, newsletters, receipts)
   - Inside each item: bold title, muted sender + date line (12px), summary paragraph (13px).

3. TODAY'S SCHEDULE section:
   - Two-column layout: left col (110px) shows time in bold #3d7ab5, right col shows
     event name + location + key attendees.

4. EXTERNAL MEETING INTEL section:
   - Card per external attendee (bg #f5f6f8, border-radius 8px).
   - Name + role, blue subtitle with title/company/email, bio paragraph.
   - Only include attendees who are external (not @grahamcos.com).

5. FIVE-DAY LOOK AHEAD section:
   - Group by day (bold date header). Each event on its own line: time in blue + name.
   - Flag scheduling conflicts with ⚠ and red color (#d9534f).

6. FOOTER: light bg #f5f6f8, centered 11px muted text.

All sections separated by a 1px #e8ecf2 divider row.
Use inline CSS only (email clients strip <style> tags).
"""


def generate_briefing_html(emails: list, events: list) -> str:
    today_et = datetime.datetime.now(EASTERN)
    today_label = today_et.strftime("%A, %B %-d, %Y")

    prompt = f"""You are generating the daily morning briefing email for Andre, CEO of Graham Companies.

Today is {today_label} (Eastern Time).

=== EMAILS (last 24 hours) ===
{json.dumps(emails, indent=2)}

=== CALENDAR EVENTS (next 5 days, Eastern Time) ===
{json.dumps(events, indent=2)}

=== TASK ===
Write a complete, standalone HTML email following the style guide below.
Include all four sections: Email Summary, Today's Schedule, External Meeting Intel,
and Five-Day Look Ahead.

For External Meeting Intel, research and write substantive background on each
external attendee appearing in today's meetings (anyone not from grahamcos.com).
Include their role, company, notable background, and any relevant context for Andre.

=== STYLE GUIDE ===
{STYLE_GUIDE}

Return ONLY the complete HTML document. No explanation, no markdown fencing."""

    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def main():
    token = get_access_token()
    user_email = os.environ["MS_USER_EMAIL"]

    print("Fetching emails...")
    emails = get_emails(token, user_email)
    print(f"  {len(emails)} emails retrieved")

    print("Fetching calendar events...")
    events = get_calendar_events(token, user_email)
    print(f"  {len(events)} events retrieved")

    print("Generating briefing with Claude...")
    html = generate_briefing_html(emails, events)

    with open("briefing_body.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Briefing saved to briefing_body.html")


if __name__ == "__main__":
    main()
