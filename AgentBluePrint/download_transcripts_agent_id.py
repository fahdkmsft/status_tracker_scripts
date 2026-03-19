import argparse
import os
import requests
import msal

# ================= DEFAULT CONFIG =================
DEFAULT_TENANT_ID = "9c23c1e3-15be-4744-a3d7-027089c33654" # agent365001
DEFAULT_CLIENT_ID = "80e89142-d630-4321-8a1e-c77f72a52220" # fahdkteamsapp

AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{}"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

SCOPES = ["https://graph.microsoft.com/.default"]
# ==================================================


def get_access_token(tenant_id, client_id, client_secret):
    authority = AUTHORITY_TEMPLATE.format(tenant_id)

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=SCOPES)

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


def get_meeting_transcripts(token, user_id):
    headers = {"Authorization": f"Bearer {token}"}

    transcripts = []

    # List online meetings for the user
    meetings_url = f"{GRAPH_BASE}/users/{user_id}/onlineMeetings"

    print("Fetching online meetings...")
    while meetings_url:
        resp = requests.get(meetings_url, headers=headers)

        if not resp.ok:
            log_api_error(resp)
            resp.raise_for_status()

        data = resp.json()
        meetings = data.get("value", [])
        print(f"  Found {len(meetings)} online meetings")

        for meeting in meetings:
            meeting_id = meeting["id"]
            subject = meeting.get("subject", "No subject")
            start_time = meeting.get("startDateTime", "")

            # Get transcripts for this meeting
            transcripts_url = f"{GRAPH_BASE}/users/{user_id}/onlineMeetings/{meeting_id}/transcripts"
            resp = requests.get(transcripts_url, headers=headers)

            if resp.status_code in (403, 404):
                continue

            if not resp.ok:
                log_api_error(resp)
                continue

            transcript_data = resp.json()
            for transcript in transcript_data.get("value", []):
                transcript["_meeting_subject"] = subject
                transcript["_meeting_start"] = start_time
                transcripts.append((meeting_id, transcript))

        meetings_url = data.get("@odata.nextLink")

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


def download_transcript(token, user_id, meeting_id, transcript, output_dir, subject, meeting_time):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "text/vtt",  # Request VTT format
    }

    transcript_id = transcript["id"]
    
    content_url = (
        f"{GRAPH_BASE}/users/{user_id}/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content"
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
    parser.add_argument("--client-secret", required=True, help="Client secret for app authentication")
    parser.add_argument("--user-id", required=True, help="User ID or UPN to fetch transcripts for")
    parser.add_argument("--output-dir", default="transcripts")

    args = parser.parse_args()

    token = get_access_token(args.tenant_id, args.client_id, args.client_secret)
    transcripts = get_meeting_transcripts(token, args.user_id)

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
                args.user_id,
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
