"""
fetch_meeting_info.py

Retrieves joinMeetingId + passcode for a meeting that was already scheduled
via Graph, so you don't have to copy them by hand from an invite.

Requires the same organizer-scoped access as creating the meeting did:
OnlineMeetings.Read.All (or ReadWrite.All) application permission, plus an
Application Access Policy granted to the organizer -- if you were able to
create the meeting via Graph, this is already in place.

Required environment variables:
  TENANT_ID, CLIENT_ID, CLIENT_SECRET   -- same as initiate_call.py
  ORGANIZER_USER_ID                     -- the organizer's AAD object ID
  ONLINE_MEETING_ID                     -- the "id" field from meeting creation,
                                           OR set JOIN_WEB_URL instead if that's
                                           the only thing you saved
"""

import os
import requests
import msal

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
ORGANIZER_USER_ID = os.environ["ORGANIZER_USER_ID"]

ONLINE_MEETING_ID = os.environ.get("ONLINE_MEETING_ID")
JOIN_WEB_URL = os.environ.get("JOIN_WEB_URL")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description')}")
    return result["access_token"]


def fetch_meeting(access_token: str) -> dict:
    headers = {"Authorization": f"Bearer {access_token}"}

    if ONLINE_MEETING_ID:
        url = f"{GRAPH_BASE}/users/{ORGANIZER_USER_ID}/onlineMeetings/{ONLINE_MEETING_ID}"
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        return resp.json()

    if JOIN_WEB_URL:
        url = f"{GRAPH_BASE}/users/{ORGANIZER_USER_ID}/onlineMeetings"
        resp = requests.get(url, headers=headers, params={"$filter": f"JoinWebUrl eq '{JOIN_WEB_URL}'"}, timeout=15)
        resp.raise_for_status()
        results = resp.json().get("value", [])
        if not results:
            raise RuntimeError("No meeting found matching that JoinWebUrl.")
        return results[0]

    raise RuntimeError("Set either ONLINE_MEETING_ID or JOIN_WEB_URL.")


def main():
    token = get_access_token()
    meeting = fetch_meeting(token)

    settings = meeting.get("joinMeetingIdSettings", {})
    join_meeting_id = settings.get("joinMeetingId")
    passcode = settings.get("passcode")

    print(f"JOIN_MEETING_ID={join_meeting_id}")
    print(f"MEETING_PASSCODE={passcode}")
    print("\nSet these as environment variables before running initiate_call.py.")


if __name__ == "__main__":
    main()
