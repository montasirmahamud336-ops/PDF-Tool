import asyncio
from telethon import TelegramClient, functions

name = "random_name"
api_id = 12345678
api_hash = "random_api_hash"

invite_link = "https://t.me/+JMAPFmH2Li04OTdl"

async def main():
    async with TelegramClient(name, api_id, api_hash) as client:
        result = await client(functions.messages.ImportChatInviteRequest(invite_link))
        print(result.stringify())

        group_info = await client(functions.groups.GetFullChatRequest(result.chat_id))
        linked_chat_id = group_info.full_chat.linked_chat_id

        if linked_chat_id:
            join_result = await client(functions.channels.JoinChannelRequest(linked_chat_id))
            print(join_result.stringify())