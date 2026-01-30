import argparse
import os
import re
import requests
import msal
from datetime import datetime, timedelta, timezone

# ================= DEFAULT CONFIG =================
DEFAULT_TENANT_ID = "adfa4542-3e1e-46f5-9c70-3df0b15b3f6c"  # a365preview001
DEFAULT_CLIENT_ID = "f3238143-69b7-4aa8-8da6-dbeef8b7b1dc"  # fahdkteamsapp

AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{}"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

SCOPES = [
    "User.Read",
    "Chat.Read",
    "Chat.ReadBasic",
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


def list_chats(token):
    headers = {"Authorization": f"Bearer {token}"}

    url = f"{GRAPH_BASE}/me/chats?$expand=members&$top=50"
    chats = []

    print("Fetching chats...")
    while url:
        resp = requests.get(url, headers=headers)
        
        if not resp.ok:
            log_api_error(resp)
            resp.raise_for_status()
        
        data = resp.json()
        chats.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return chats


def get_chat_display_name(chat):
    """Get a display name for the chat."""
    topic = chat.get("topic")
    if topic:
        return topic
    
    chat_type = chat.get("chatType", "unknown")
    
    # For one-on-one or group chats without topic, use member names
    members = chat.get("members", [])
    if members:
        member_names = []
        for member in members:
            display_name = member.get("displayName")
            if display_name:
                member_names.append(display_name)
        if member_names:
            return f"{chat_type}: {', '.join(member_names[:3])}"
    
    return f"{chat_type} chat"


def get_chat_messages(token, chat_id, target_date):
    """Get messages for a chat on a specific date."""
    headers = {"Authorization": f"Bearer {token}"}
    
    # Calculate date range for the target date
    start_of_day = datetime.combine(target_date, datetime.min.time()).replace(tzinfo=timezone.utc)
    end_of_day = start_of_day + timedelta(days=1)
    
    # Graph API doesn't support filtering chat messages by createdDateTime
    # Fetch messages and filter client-side
    url = f"{GRAPH_BASE}/me/chats/{chat_id}/messages?$top=50"
    
    messages = []
    found_older = False
    
    while url and not found_older:
        resp = requests.get(url, headers=headers)
        
        if not resp.ok:
            log_api_error(resp)
            # Return empty if we can't access messages
            if resp.status_code in (403, 404):
                return []
            resp.raise_for_status()
        
        data = resp.json()
        
        for msg in data.get("value", []):
            created_str = msg.get("createdDateTime", "")
            if not created_str:
                continue
            
            try:
                # Parse the datetime string
                created_dt = datetime.fromisoformat(created_str.replace("Z", "+00:00"))
                
                # Check if message is within target date
                if start_of_day <= created_dt < end_of_day:
                    messages.append(msg)
                elif created_dt < start_of_day:
                    # Messages are returned newest first, so if we hit older messages, stop
                    found_older = True
                    break
            except ValueError:
                continue
        
        url = data.get("@odata.nextLink")
    
    return messages


def sanitize_filename(name):
    """Remove or replace characters that are invalid in filenames."""
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Replace spaces with hyphens
    name = name.replace(' ', '-')
    name = name.strip(' .')
    return name[:100] if len(name) > 100 else name


def format_message(msg):
    """Format a message for output."""
    sender = msg.get("from", {})
    user = sender.get("user", {}) if sender else {}
    display_name = user.get("displayName", "Unknown") if user else "System"
    
    created = msg.get("createdDateTime", "")
    body = msg.get("body", {})
    content = body.get("content", "")
    content_type = body.get("contentType", "text")
    
    # Strip HTML tags if content is HTML
    if content_type == "html":
        content = re.sub(r'<[^>]+>', '', content)
    
    # Clean up whitespace
    content = content.strip()
    
    return f"[{created}] {display_name}: {content}"


def download_chat_messages(token, chat, target_date, output_dir):
    """Download messages for a chat on a specific date."""
    chat_id = chat["id"]
    chat_name = get_chat_display_name(chat)
    
    print(f"Fetching messages for '{chat_name}' on {target_date}...")
    
    messages = get_chat_messages(token, chat_id, target_date)
    
    if not messages:
        print(f"  No messages found for {target_date}")
        return False
    
    os.makedirs(output_dir, exist_ok=True)
    
    # Create filename from chat name and date
    safe_name = sanitize_filename(chat_name)
    date_str = target_date.strftime("%Y-%m-%d")
    filename = os.path.join(output_dir, f"{safe_name}_{date_str}.txt")
    
    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"Chat: {chat_name}\n")
        f.write(f"Date: {date_str}\n")
        f.write(f"Messages: {len(messages)}\n")
        f.write("=" * 50 + "\n\n")
        
        for msg in sorted(messages, key=lambda m: m.get("createdDateTime", "")):
            formatted = format_message(msg)
            if formatted.strip():
                f.write(formatted + "\n\n")
    
    print(f"  Downloaded {len(messages)} messages → {filename}")
    return True


def prompt_yes_no(prompt):
    """Prompt the user for a yes/no answer."""
    while True:
        response = input(f"{prompt} (y/n): ").strip().lower()
        if response in ("y", "yes"):
            return True
        if response in ("n", "no"):
            return False
        print("Please enter 'y' or 'n'")


def prompt_date(prompt):
    """Prompt the user for a date."""
    while True:
        response = input(f"{prompt} (YYYY-MM-DD, or 'today'/'yesterday'): ").strip().lower()
        
        today = datetime.now().date()
        
        if response == "today":
            return today
        if response == "yesterday":
            return today - timedelta(days=1)
        
        try:
            return datetime.strptime(response, "%Y-%m-%d").date()
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD, 'today', or 'yesterday'")


def main():
    parser = argparse.ArgumentParser(
        description="Download Microsoft Teams chat messages"
    )
    parser.add_argument("--tenant-id", default=DEFAULT_TENANT_ID)
    parser.add_argument("--client-id", default=DEFAULT_CLIENT_ID)
    parser.add_argument("--output-dir", default="chats")
    parser.add_argument(
        "--date",
        help="Date to download messages for (YYYY-MM-DD, 'today', or 'yesterday')"
    )

    args = parser.parse_args()

    # Get the target date
    if args.date:
        if args.date.lower() == "today":
            target_date = datetime.now().date()
        elif args.date.lower() == "yesterday":
            target_date = datetime.now().date() - timedelta(days=1)
        else:
            try:
                target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
            except ValueError:
                print("Invalid date format. Use YYYY-MM-DD, 'today', or 'yesterday'")
                return
    else:
        target_date = prompt_date("Which date to download messages for?")

    print(f"\nTarget date: {target_date}\n")

    token = get_access_token(args.tenant_id, args.client_id)
    chats = list_chats(token)

    print(f"\nFound {len(chats)} chats\n")

    downloaded_count = 0
    skipped_count = 0

    for i, chat in enumerate(chats, 1):
        chat_name = get_chat_display_name(chat)
        chat_type = chat.get("chatType", "unknown")
        
        print(f"\n--- Chat {i}/{len(chats)} ---")
        print(f"  Name: {chat_name}")
        print(f"  Type: {chat_type}")
        
        if prompt_yes_no("Download messages for this chat?"):
            if download_chat_messages(token, chat, target_date, args.output_dir):
                downloaded_count += 1
            else:
                skipped_count += 1
        else:
            print("  Skipped.")
            skipped_count += 1

    print(f"\n=== Summary ===")
    print(f"Downloaded: {downloaded_count}")
    print(f"Skipped:    {skipped_count}")


if __name__ == "__main__":
    main()
