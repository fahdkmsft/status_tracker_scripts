import json
import subprocess
import requests
import shutil


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_graph_token_from_az():
    """
    Calls Azure CLI to get an access token for Microsoft Graph
    """
    az_path = shutil.which("az")

    if not az_path:
        raise RuntimeError(
            "Azure CLI (az) not found in PATH. "
            "Run 'az --version' in this shell first."
        )
    
    result = subprocess.run(
        [
            az_path,
            "account",
            "get-access-token",
            "--resource",
            "https://graph.microsoft.com",
            "--output",
            "json",
        ],
        capture_output=True,
        text=True,
        check=True,
    )

    token_data = json.loads(result.stdout)
    return token_data["accessToken"]


def list_chats(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = f"{GRAPH_BASE}/me/chats"
    chats = []

    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()

        chats.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return chats


def main():
    token = get_graph_token_from_az()
    chats = list_chats(token)

    print(f"\nFound {len(chats)} chats:\n")

    for chat in chats:
        print("===================================")
        print(f"Chat ID : {chat['id']}")
        print(f"Type    : {chat.get('chatType')}")
        print(f"Topic   : {chat.get('topic')}")
        print("===================================\n")


if __name__ == "__main__":
    main()
