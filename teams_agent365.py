import asyncio
import json
from azure.identity import AzureCliCredential
from mcp.client.sse import sse_client


MCP_ENDPOINT = "https://agents.microsoft.com/mcp/teams"


class TeamsMCPExporter:

    def __init__(self):
        self.credential = AzureCliCredential()
        self.token = None

    def authenticate(self):
        print("Getting Azure CLI token...")
        token = self.credential.get_token(
            "https://graph.microsoft.com/.default"
        )
        self.token = token.token
        print("Authentication successful")

    async def list_chats(self):

        async with sse_client(
            MCP_ENDPOINT,
            headers={"Authorization": f"Bearer {self.token}"}
        ) as client:

            result = await client.call_tool(
                "microsoft_teams_list_chats",
                {}
            )

            return result.get("chats", [])

    async def list_messages(self, chat_id):

        async with sse_client(
            MCP_ENDPOINT,
            headers={"Authorization": f"Bearer {self.token}"}
        ) as client:

            result = await client.call_tool(
                "microsoft_teams_list_messages",
                {
                    "chat_id": chat_id
                }
            )

            return result.get("messages", [])

    async def export_all(self):

        chats = await self.list_chats()

        print(f"Found {len(chats)} chats")

        export = []

        for chat in chats:

            chat_id = chat["id"]
            topic = chat.get("topic", "")

            print(f"Fetching messages for chat: {topic}")

            messages = await self.list_messages(chat_id)

            export.append({
                "chat_id": chat_id,
                "topic": topic,
                "messages": messages
            })

        return export


async def main():

    exporter = TeamsMCPExporter()

    exporter.authenticate()

    data = await exporter.export_all()

    with open("teams_export.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print("Export complete → teams_export.json")


if __name__ == "__main__":
    asyncio.run(main())
