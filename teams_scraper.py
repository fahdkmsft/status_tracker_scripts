from playwright.sync_api import sync_playwright
from pathlib import Path
from datetime import datetime
import json
import time
import os

# Use Edge's default user data directory to access existing profiles
EDGE_USER_DATA = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Edge", "User Data")
# Change this to "Profile 1" if your work account is in the second Edge profile
EDGE_PROFILE = "Default"  # or "Profile 1" for the second profile
PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "teams-profile")
OUTPUT_FILE = "teams_output.json"


def get_user_date():
    """Prompt user for a specific date to filter data."""
    while True:
        date_str = input("Enter date to download data for (YYYY-MM-DD): ").strip()
        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d")
            return target_date
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")


def display_and_select(items, item_type):
    """Display a numbered list and let user select items."""
    if not items:
        print(f"No {item_type} found.")
        return []

    print(f"\nAvailable {item_type}:")
    for i, item in enumerate(items, 1):
        print(f"  {i}. {item}")

    print(f"\nEnter {item_type} numbers to download (comma-separated), 'all' for all, or 'none' to skip:")
    selection = input("> ").strip().lower()

    if selection == 'none' or selection == '':
        return []
    if selection == 'all':
        return list(range(len(items)))

    try:
        indices = [int(x.strip()) - 1 for x in selection.split(',')]
        # Validate indices
        valid_indices = [i for i in indices if 0 <= i < len(items)]
        return valid_indices
    except ValueError:
        print("Invalid selection. Skipping.")
        return []


def scrape_teams():
    # Get target date from user
    target_date = get_user_date()
    print(f"\nFiltering data for: {target_date.strftime('%Y-%m-%d')}")
    
    print("\nIMPORTANT: Please close ALL Edge browser windows before continuing.")
    print("Press Enter when ready...")
    input()

    with sync_playwright() as p:
        # Launch Edge using your existing profile (avoids profile picker)
        context = p.chromium.launch_persistent_context(
            EDGE_USER_DATA,
            headless=False,
            channel="msedge",
            args=[
                f"--profile-directory={EDGE_PROFILE}",
                "--start-maximized",
            ],
            slow_mo=100,
        )
        
        page = context.pages[0] if context.pages else context.new_page()

        print("Opening Teams...")
        page.goto("https://teams.microsoft.com")

        # Wait for Teams to fully load
        print("Waiting for Teams to load (this may take a moment)...")
        page.wait_for_url("**/teams.cloud.microsoft/**", timeout=180_000)
        page.wait_for_timeout(10000)  # Give Teams 10 seconds to stabilize
        print("Teams loaded!")

        data = {
            "date": target_date.strftime("%Y-%m-%d"),
            "chats": [],
            "transcripts": []
        }

        # ---------------------------
        # READ CHAT MESSAGES
        # ---------------------------
        print("\n" + "="*50)
        print("CHAT MESSAGES")
        print("="*50)
        
        while True:
            print("\nTo download a chat:")
            print("  1. In the Teams window, open the chat you want")
            print("  2. Come back here and press Enter")
            print("  3. Or type 'done' to move to transcripts")
            
            user_input = input("\n> ").strip().lower()
            if user_input == 'done':
                break
            
            # User pressed Enter - try to get chat from current view
            print("Downloading chat from current view...")
            
            # Try to get the chat name from the header
            chat_name = "Unknown Chat"
            try:
                title_element = page.locator('[data-tid*="chat-header"], [data-tid*="conversation-header"], h1, h2').first
                if title_element.count() > 0:
                    chat_name = title_element.inner_text().strip().split('\n')[0][:60]
            except:
                pass
            
            # Scroll up to load older messages
            print("  Scrolling up to load older messages...")
            try:
                message_container = page.locator('[data-tid*="message-pane"], [role="main"], .ts-message-list-container').first
                for _ in range(10):  # Scroll up multiple times
                    message_container.evaluate("el => el.scrollTop = 0")
                    page.wait_for_timeout(1000)
            except Exception as e:
                print(f"  Warning: Could not scroll: {e}")
            
            # Get messages
            messages = page.locator(
                '[data-tid*="message"], [data-tid*="messageBody"], .message-body-content'
            ).all_inner_texts()
            
            if not messages:
                print("  Could not find messages in current view.")
                print("  Make sure a chat is open and visible.")
                continue
            
            # Ask user for a name for this chat
            name_input = input(f"  Enter name for this chat [{chat_name}]: ").strip()
            if name_input:
                chat_name = name_input
            
            data["chats"].append({
                "chat_name": chat_name,
                "messages": messages
            })
            print(f"  ✓ Saved chat: {chat_name} ({len(messages)} messages)")

        # ---------------------------
        # READ MEETING TRANSCRIPTS
        # ---------------------------
        print("\n" + "="*50)
        print("MEETING TRANSCRIPTS")
        print("="*50)
        
        while True:
            print("\nTo download a meeting transcript:")
            print("  1. In the Teams window, open the meeting chat you want")
            print("  2. Come back here and press Enter")
            print("  3. Or type 'done' to finish and save")
            
            user_input = input("\n> ").strip().lower()
            if user_input == 'done':
                break
            
            # User pressed Enter - try to get transcript from current view
            print("Looking for transcript in the current meeting chat...")
            
            # Look for the Recap/Transcript tab or button
            recap_button = page.locator('button:has-text("Recap"), [aria-label*="Recap"], [data-tid*="recap"]').first
            if recap_button.count() > 0:
                print("  Found Recap button, clicking...")
                recap_button.click()
                page.wait_for_timeout(2000)
            
            # Look for transcript tab/button
            transcript_tab = page.locator('button:has-text("Transcript"), [aria-label*="Transcript"], [data-tid*="transcript"]').first
            if transcript_tab.count() > 0:
                print("  Found Transcript tab, clicking...")
                transcript_tab.click()
                page.wait_for_timeout(3000)
            
            # Try to get the meeting name from the header
            meeting_name = "Unknown Meeting"
            try:
                # Try various selectors for the meeting/chat title
                title_element = page.locator('[data-tid*="header"] h1, [data-tid*="header"] h2, .conversation-title, [data-tid*="chat-header"]').first
                if title_element.count() > 0:
                    meeting_name = title_element.inner_text().strip()[:60]
            except:
                pass
            
            # Look for transcript content
            transcript_lines = []
            
            # Try various transcript selectors
            transcript_selectors = [
                '[data-tid*="transcript-segment"]',
                '[data-tid*="transcript-line"]', 
                '[data-tid*="transcript-item"]',
                '.transcript-segment',
                '.transcript-line',
                '[role="listitem"]'  # Transcript items might be list items
            ]
            
            for selector in transcript_selectors:
                items = page.locator(selector)
                if items.count() > 0:
                    print(f"  Found {items.count()} transcript items with selector: {selector}")
                    transcript_lines = items.all_inner_texts()
                    break
            
            if not transcript_lines:
                print("  Could not find transcript content.")
                print("  Make sure the transcript tab is open and visible.")
                print("  Press F12 to inspect and find the right selector.")
                continue
            
            # Ask user for a name for this transcript
            name_input = input(f"  Enter name for this transcript [{meeting_name}]: ").strip()
            if name_input:
                meeting_name = name_input
            
            data["transcripts"].append({
                "meeting_name": meeting_name,
                "lines": transcript_lines
            })
            print(f"  ✓ Saved transcript: {meeting_name} ({len(transcript_lines)} lines)")

        # Save with date in filename
        output_file = f"teams_output_{target_date.strftime('%Y-%m-%d')}.json"
        Path(output_file).write_text(
            json.dumps(data, indent=2),
            encoding="utf-8"
        )

        print(f"\nSaved {output_file}")
        context.close()


if __name__ == "__main__":
    scrape_teams()
