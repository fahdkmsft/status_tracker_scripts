import argparse
import os
import requests
import msal

# ================= DEFAULT CONFIG =================
DEFAULT_TENANT_ID = "adfa4542-3e1e-46f5-9c70-3df0b15b3f6c" # a365preview001
DEFAULT_CLIENT_ID = "f3238143-69b7-4aa8-8da6-dbeef8b7b1dc" # fahdkteamsapp

AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{}"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

SCOPES = [
    "User.Read",
    "Calendars.Read",
    "OnlineMeetings.Read",
    "OnlineMeetingTranscript.Read.All",
]
# ==================================================


def get_access_token(tenant_id, client_id):
    authority = AUTHORITY_TEMPLATE.format(tenant_id)

    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
    )

    result = app.acquire_token_interactive(
        scopes=SCOPES,
        prompt="select_account",
    )

    if "access_token" not in result:
        raise RuntimeError(f"Authentication failed: {result}")

    return result["access_token"]


def log_api_error(resp):
    """Log API error details from response."""
    print(f"API Error: {resp.status_code} {resp.reason}")
    try:
        error_data = resp.json()
        if "error" in error_data:
            error = error_data["error"]
            print(f"  Code: {error.get('code', 'N/A')}")
            print(f"  Message: {error.get('message', 'N/A')}")
        else:
            print(f"  Response: {resp.text}")
    except Exception:
        print(f"  Response: {resp.text}")


def get_meeting_transcripts(token):
    headers = {"Authorization": f"Bearer {token}"}

    # Get recent calendar events (can't filter by isOnlineMeeting)
    from datetime import datetime, timedelta, timezone
    end_date = datetime.now(timezone.utc)
    start_date = end_date - timedelta(days=90)
    
    # Use calendarView to get events in a date range
    start_str = start_date.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end_date.strftime("%Y-%m-%dT%H:%M:%SZ")
    
    events_url = (
        f"{GRAPH_BASE}/me/calendarView"
        f"?startDateTime={start_str}"
        f"&endDateTime={end_str}"
        f"&$select=id,subject,start,end,onlineMeeting,isOnlineMeeting"
        f"&$orderby=start/dateTime desc"
        f"&$top=100"
    )
    
    transcripts = []
    seen_meeting_ids = set()

    print("Fetching calendar events from the last 90 days...")
    while events_url:
        resp = requests.get(events_url, headers=headers)
        
        if not resp.ok:
            log_api_error(resp)
            resp.raise_for_status()
        
        data = resp.json()
        events = data.get("value", [])
        online_events = [e for e in events if e.get("isOnlineMeeting") or e.get("onlineMeeting")]
        print(f"  Found {len(events)} events, {len(online_events)} are online meetings")

        for event in online_events:
            online_meeting = event.get("onlineMeeting")
            if not online_meeting:
                continue
            
            join_url = online_meeting.get("joinUrl")
            if not join_url:
                continue
            
            subject = event.get("subject", "No subject")
            start_time = event.get("start", {}).get("dateTime", "")
            
            # Get the online meeting by join URL to get its ID
            import urllib.parse
            encoded_url = urllib.parse.quote(join_url, safe='')
            filter_url = (
                f"{GRAPH_BASE}/me/onlineMeetings"
                f"?$filter=joinWebUrl eq '{join_url}'"
            )
            
            resp = requests.get(filter_url, headers=headers)
            
            if not resp.ok:
                # Skip if we can't get the meeting
                continue
            
            meetings_data = resp.json()
            meetings = meetings_data.get("value", [])
            
            if not meetings:
                continue
            
            meeting = meetings[0]
            meeting_id = meeting["id"]
            
            if meeting_id in seen_meeting_ids:
                continue
            seen_meeting_ids.add(meeting_id)
            
            # Get transcripts for this meeting
            transcripts_url = f"{GRAPH_BASE}/me/onlineMeetings/{meeting_id}/transcripts"
            resp = requests.get(transcripts_url, headers=headers)
            
            if resp.status_code in (403, 404):
                # No transcripts or no permission for this meeting
                continue
            
            if not resp.ok:
                log_api_error(resp)
                continue
            
            transcript_data = resp.json()
            for transcript in transcript_data.get("value", []):
                transcript["_meeting_subject"] = subject
                transcript["_meeting_start"] = start_time
                transcripts.append((meeting_id, transcript))

        events_url = data.get("@odata.nextLink")

    return transcripts


import re


def sanitize_filename(name):
    """Remove or replace characters that are invalid in filenames."""
    # Replace invalid characters with underscores
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Replace spaces with hyphens
    name = name.replace(' ', '-')
    # Remove leading/trailing spaces and dots
    name = name.strip(' .')
    # Limit length
    return name[:100] if len(name) > 100 else name


def download_transcript(token, meeting_id, transcript, output_dir, subject, meeting_time):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "text/vtt",  # Request VTT format
    }

    transcript_id = transcript["id"]
    
    # Use /me/onlineMeetings endpoint for delegated permissions
    content_url = (
        f"{GRAPH_BASE}/me/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content"
    )

    resp = requests.get(content_url, headers=headers)
    if not resp.ok:
        print(f"Failed to download transcript:")
        print(f"  URL: {content_url}")
        log_api_error(resp)
        resp.raise_for_status()

    os.makedirs(output_dir, exist_ok=True)
    
    # Create filename from subject and meeting time
    safe_subject = sanitize_filename(subject)
    # Parse and format the timestamp (e.g., "2026-01-26T21:00:00.0000000" -> "2026-01-26_2100")
    timestamp = ""
    if meeting_time and meeting_time != "Unknown":
        try:
            # Extract date and time parts
            dt_part = meeting_time.split(".")[0]  # Remove fractional seconds
            timestamp = dt_part.replace(":", "").replace("T", "_")[:15]  # "2026-01-26_2100"
        except Exception:
            timestamp = "unknown_time"
    
    filename = os.path.join(
        output_dir, f"{safe_subject}_{timestamp}.txt"
    )

    with open(filename, "wb") as f:
        f.write(resp.content)

    print(f"Downloaded transcript → {filename}")


def prompt_yes_no(prompt):
    """Prompt the user for a yes/no answer."""
    while True:
        response = input(f"{prompt} (y/n): ").strip().lower()
        if response in ("y", "yes"):
            return True
        if response in ("n", "no"):
            return False
        print("Please enter 'y' or 'n'")


def main():
    parser = argparse.ArgumentParser(
        description="Download Microsoft Teams meeting transcripts"
    )
    parser.add_argument("--tenant-id", default=DEFAULT_TENANT_ID)
    parser.add_argument("--client-id", default=DEFAULT_CLIENT_ID)
    parser.add_argument("--output-dir", default="transcripts")

    args = parser.parse_args()

    token = get_access_token(args.tenant_id, args.client_id)
    transcripts = get_meeting_transcripts(token)

    print(f"Found {len(transcripts)} transcripts")

    downloaded_count = 0
    skipped_count = 0

    for meeting_id, transcript in transcripts:
        transcript_id = transcript["id"]
        created_time = transcript.get("createdDateTime", "Unknown time")
        subject = transcript.get("_meeting_subject", "Unknown")
        meeting_start = transcript.get("_meeting_start", "Unknown")
        
        print(f"\n--- Transcript ---")
        print(f"  Subject:       {subject}")
        print(f"  Meeting Time:  {meeting_start}")
        print(f"  Transcript ID: {transcript_id}")
        print(f"  Created:       {created_time}")
        
        if prompt_yes_no("Download this transcript?"):
            download_transcript(
                token,
                meeting_id,
                transcript,
                args.output_dir,
                subject,
                meeting_start,
            )
            downloaded_count += 1
        else:
            print("Skipped.")
            skipped_count += 1

    print(f"\n=== Summary ===")
    print(f"Downloaded: {downloaded_count}")
    print(f"Skipped:    {skipped_count}")


if __name__ == "__main__":
    main()
