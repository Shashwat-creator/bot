"""
initiate_call.py

Runs on your private server. Places an outbound call to Microsoft Graph that
joins an existing Teams meeting as a guest, using the meeting's joinMeetingId
+ passcode (Path A from our plan -- no Application Access Policy needed).

This only makes OUTBOUND HTTPS requests, so no public/inbound endpoint is
needed on this machine. The CALLBACK_URI below points at the Azure stub
(the other file), not at this server.

Required environment variables:
  TENANT_ID        Azure AD tenant ID
  CLIENT_ID        Azure AD app (client) ID
  CLIENT_SECRET    Azure AD app client secret
  CALLBACK_URI     Public HTTPS URL of the Azure webhook stub,
                   e.g. https://<your-app>.azurewebsites.net/api/calling
  JOIN_MEETING_ID  Numeric meeting ID from the calendar invite
  MEETING_PASSCODE Optional -- omit if the meeting has no passcode
"""

import os
import time
import requests
import msal

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
CALLBACK_URI = os.environ["CALLBACK_URI"]
JOIN_MEETING_ID = os.environ["JOIN_MEETING_ID"]
MEETING_PASSCODE = os.environ.get("MEETING_PASSCODE")

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


def join_meeting(access_token: str) -> str:
    """POSTs to /communications/calls to join the meeting as a guest. Returns the call id."""
    meeting_info = {
        "@odata.type": "#microsoft.graph.joinMeetingIdMeetingInfo",
        "joinMeetingId": JOIN_MEETING_ID,
    }
    if MEETING_PASSCODE:
        meeting_info["passcode"] = MEETING_PASSCODE

    payload = {
        "@odata.type": "#microsoft.graph.call",
        "callbackUri": CALLBACK_URI,
        "requestedModalities": ["audio"],
        "mediaConfig": {"@odata.type": "#microsoft.graph.serviceHostedMediaConfig"},
        "meetingInfo": meeting_info,
        "tenantId": TENANT_ID,
    }

    resp = requests.post(
        f"{GRAPH_BASE}/communications/calls",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )

    if resp.status_code not in (200, 201, 202):
        raise RuntimeError(f"Join call failed: {resp.status_code} {resp.text}")

    call_id = resp.json()["id"]
    print(f"Call created: {call_id}")
    return call_id


def poll_call_state(access_token: str, call_id: str,
                     timeout_seconds: int = 120, interval_seconds: int = 5) -> str:
    """
    Polls the call resource directly instead of waiting on a webhook -- this is
    the "no relay needed for testing" path. Returns the final observed state.
    """
    deadline = time.time() + timeout_seconds
    last_state = None

    while time.time() < deadline:
        resp = requests.get(
            f"{GRAPH_BASE}/communications/calls/{call_id}",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=15,
        )

        if resp.status_code == 404:
            print("Call no longer exists (likely terminated and cleaned up).")
            return "terminated"

        resp.raise_for_status()
        state = resp.json().get("state")

        if state != last_state:
            print(f"Call state: {state}")
            last_state = state

        if state in ("established", "terminated"):
            return state

        time.sleep(interval_seconds)

    print("Timed out waiting for the call to reach a terminal state.")
    return last_state


def main():
    token = get_access_token()
    call_id = join_meeting(token)
    final_state = poll_call_state(token, call_id)
    print(f"Finished with state: {final_state}")


if __name__ == "__main__":
    main()
