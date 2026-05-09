    """
Sends the generated briefing HTML via Microsoft Graph sendMail.
"""
import json
import os
import datetime
import zoneinfo
import requests

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


def main():
    today_label = datetime.datetime.now(EASTERN).strftime("%A, %B %-d, %Y")
    subject = f"Morning Briefing — {today_label}"

    with open("briefing_body.html", "r", encoding="utf-8") as f:
        html = f.read()

    token = get_access_token()
    user_email = os.environ["MS_USER_EMAIL"]

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [
                {"emailAddress": {"address": user_email}}
            ],
        },
        "saveToSentItems": False,
    }

    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
    )
    resp.raise_for_status()
    print(f"SUCCESS: \"{subject}\" sent to {user_email}")


if __name__ == "__main__":
    main()
