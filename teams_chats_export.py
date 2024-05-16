import argparse
import asyncio
import base64
from functools import cache
import glob
import json
import os
import pprint
import re
import shutil
from typing import Dict, Optional

from azure.identity import InteractiveBrowserCredential
import dateparser
from jinja2 import Environment, FileSystemLoader
from kiota_abstractions.native_response_handler import NativeResponseHandler
from kiota_http.middleware.options import ResponseHandlerOption
from msgraph import GraphServiceClient
from msgraph.generated.chats.chats_request_builder import ChatsRequestBuilder
from msgraph.generated.chats.item.messages.messages_request_builder import MessagesRequestBuilder
import pytz

# this is the universal client id (aka application id) used by Microsoft's "PnP PowerShell".
# users typically are granted access to this in every organization and lets us avoid having
# to create an Azure Application and grant it permissions for use with the MS Graph API.
# see https://pnp.github.io/powershell/cmdlets/Request-PnPAccessToken.html
pnp_management_shell_client_id = "31359c7f-bd7e-475c-86db-fdb8c937548e"

filename_size_limit = 255


def makedir(path):
    """basically mkdir -p"""
    if not os.path.exists(path):
        os.makedirs(os.path.join(path), exist_ok=True)


@cache
def get_jinja_env():
    jinja_env = Environment(loader=FileSystemLoader("templates"))
    jinja_env.filters["localdt"] = localdt
    return jinja_env


def localdt(value: str, format="%m/%d/%Y %I:%M %p %Z"):
    """parse a date string into a datetime object, localize it, and format it for display"""
    tz = pytz.timezone("America/Los_Angeles")
    dt = dateparser.parse(value)
    local_dt = dt.astimezone(tz)
    return local_dt.strftime(format)


def get_member_list(chat: dict):
    """return a sorted comma-separated list of chat members"""
    members = [
        m["displayName"] if m["displayName"] else "No Name" for m in chat["members"]
    ]
    return ", ".join(sorted(members))


def get_chat_name(chat: dict):
    """get a "name" for the chat: either its topic or a comma-separated list of members"""
    if chat["topic"]:
        name = chat["topic"].replace(os.path.sep, "_")
    else:
        name = get_member_list(chat)
    return name


def get_hosted_content_filename(msg_id, hosted_content_id):
    """
    return a base filename for the msg_id + hosted_content_id,
    truncating if necessary to keep it under the filename size limit
    """
    filename = f"hosted_content_{msg_id}_{hosted_content_id}"
    return filename[0:filename_size_limit]


def get_hosted_content_id(attachment: dict) -> str:
    """extract the hosted_content_id from the Attachment dict record"""
    # it's stupid that I have to parse this. codeSnippetUrl already is the complete URL
    # but I can't figure out how to make a request to it directly using the client object
    content = json.loads(attachment["content"])
    hosted_content_id = content["codeSnippetUrl"].split("/")[-2]
    return hosted_content_id


async def fetch_all_for_request(getable, request_config):
    """
    returns an iterator over the dict records returned from a request

    getable = an MS Graph API object with a get() method.
    request_config = request configuration object to pass to get()
    """
    results = None
    getable_ = getable
    while getable_:
        if results:
            if "@odata.nextLink" in results:
                getable_ = getable.with_url(results["@odata.nextLink"])
            else:
                getable_ = None
        if getable_:
            response = await getable_.get(request_configuration=request_config)
            if response:
                results = response.json()
                for result in results["value"]:
                    yield result


async def download_hosted_content(client, chat: Dict, msg: Dict, hosted_content_id: str, chat_dir: str):
    # it's happened in one case that a user doesn't have access to the hosted content
    # in a chat they're a member of. not sure how that's possible, but that's why
    # this check is here.
    try:
        result = (
            await client.chats.by_chat_id(chat["id"])
            .messages.by_chat_message_id(msg["id"])
            .hosted_contents.by_chat_message_hosted_content_id(hosted_content_id)
            .content.get()
        )
    except Exception as e:
        result = str(e)
    filename = get_hosted_content_filename(msg['id'], hosted_content_id)
    path = os.path.join(chat_dir, filename)
    with open(path, "wb") as f:
        f.write(result)


async def download_hosted_content_in_msg(client, chat: Dict, msg: Dict, chat_dir: str):
    # fetch all the "hosted contents" (inline attachments)
    for attachment in msg["attachments"]:
        if attachment["contentType"] == "application/vnd.microsoft.card.codesnippet":
            hosted_content_id = get_hosted_content_id(attachment)
            await download_hosted_content(client, chat, msg, hosted_content_id, chat_dir)

    # images are not present as attachments, just referenced in img tags
    content_type = (msg.get("body") or {}).get("contentType", "")
    content = (msg.get("body") or {}).get("content", "")
    if content_type == "html":
        for match in re.findall('src="(.+?)"', content):
            url = match
            if "https://graph.microsoft.com/v1.0/chats/" in url:
                hosted_content_id = url.split("/")[-2]
                await download_hosted_content(client, chat, msg, hosted_content_id, chat_dir)


async def download_messages(client, chat: Dict, chat_dir: str, force: bool = False):
    """
    download messages for a chat, including its 'hosted content'

    Note that msg ids are not globally unique. They're millisecond timestamps.

    the 'force' flag downloads all messages that haven't been saved yet.
    by default, only newer messages are downloaded.
    """

    async def save_msg(msg):
        with open(path, "w") as f:
            f.write(json.dumps(msg))
        await download_hosted_content_in_msg(client, chat, msg, chat_dir)

    last_msg_id = (chat["lastMessagePreview"] or {}).get("id")
    last_msg_exists = os.path.exists(os.path.join(chat_dir, f"msg_{last_msg_id}.json"))
    if force or not last_msg_id or not last_msg_exists:
        count_saved = 0
        count_updated = 0
        count_unchanged = 0
        messages_request = client.me.chats.by_chat_id(chat["id"]).messages

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=50,
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            options=[ResponseHandlerOption(NativeResponseHandler())],
            query_parameters=query_params,
        )

        async for msg in fetch_all_for_request(messages_request, request_config):
            path = os.path.join(chat_dir, f"msg_{msg['id']}.json")
            if not os.path.exists(path):
                await save_msg(msg)
                count_saved += 1
            else:
                # if incoming msg was deleted, we don't want to overwrite our file
                if not msg["deletedDateTime"]:
                    with open(path, "r") as f:
                        existing_msg = json.loads(f.read())

                    # save edited/modified msgs
                    if (
                        existing_msg["lastModifiedDateTime"] != msg["lastModifiedDateTime"]
                        or existing_msg["lastEditedDateTime"] != msg["lastEditedDateTime"]
                    ):
                        await save_msg(msg)
                        count_updated += 1
                    else:
                        count_unchanged += 1
                        # msg exists but hasn't been edited/modified, so we can stop
                        # if we're not running in force mode
                        if not force:
                            break
                else:
                    count_unchanged += 1

        output = f"  Message counts: {count_saved} saved, {count_updated} updated"
        if force:
            output += f", {count_unchanged} unchanged"
        print(output)
    else:
        print("  No new messages in the chat since last run")


async def download_chat(client, chat: Dict, data_dir: str, force: bool):
    """download a single chat and its associated data (messages, attachments)"""
    print(f"Processing chat {get_chat_name(chat)} (id {chat['id']})")

    chat_dir = os.path.join(data_dir, chat["id"])
    makedir(chat_dir)

    with open(os.path.join(data_dir, f"{chat['id']}.json"), "w") as f:
        f.write(json.dumps(chat))

    await download_messages(client, chat, chat_dir, force)


async def download_all(output_dir: str, force: bool):
    """download all chats"""
    data_dir = os.path.join(output_dir, "data")
    makedir(data_dir)

    client = get_graph_client()

    print("Opening browser window for authentication")

    query_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        expand=["members", "lastMessagePreview"], top=50
    )
    request_config = ChatsRequestBuilder.ChatsRequestBuilderGetRequestConfiguration(
        options=[ResponseHandlerOption(NativeResponseHandler())],
        query_parameters=query_params,
    )
    async for chat in fetch_all_for_request(client.me.chats, request_config):
        await download_chat(client, chat, data_dir, force)


def render_hosted_content(msg: Dict, hosted_content_id: str, chat_dir: str):
    filename = get_hosted_content_filename(msg['id'], hosted_content_id)
    path = os.path.join(chat_dir, filename)
    with open(path, "r") as f:
        data = f.read()
    return data


def render_message_body(msg: Dict, chat_dir: str, html_dir: str) -> Optional[str]:
    """render a single message body, including its attachments"""

    def get_attachment(match):
        attachment_id = match.group(1)
        attachment = [a for a in msg["attachments"] if a["id"] == attachment_id][0]
        if attachment["contentType"] == "reference":
            return f"Attachment: <a href='{attachment['contentUrl']}' data-attachment-id='{attachment['id']}'>{attachment['name']}</a><br/>"
        elif attachment["contentType"] == "messageReference":
            ref = json.loads(attachment["content"])
            return f"<blockquote class='message-reference' data-attachment-id='{attachment['id']}'>{ref['messageSender']['user']['displayName']}: {ref['messagePreview']}</blockquote>"
        elif attachment["contentType"] == "application/vnd.microsoft.card.codesnippet":
            hosted_content_id = get_hosted_content_id(attachment)
            content = render_hosted_content(msg, hosted_content_id, chat_dir)
            return f"<div class='hosted-content' data-attachment-id='{attachment['id']}' data-hosted-content-id='{hosted_content_id}'><pre><code>{content}</code></pre></div>"
        else:
            return f"Attachment (raw data): {pprint.pformat(attachment)}<br/>"

    def get_image(match):
        whole_match = match.group(0)
        url = match.group(1)
        if "https://graph.microsoft.com/v1.0/chats/" in url:
            hosted_content_id = url.split("/")[-2]
            filename = get_hosted_content_filename(msg['id'], hosted_content_id)
            with open(os.path.join(chat_dir, filename), "rb") as f:
                # TODO: not all images are actually png but this seems to work anyway
                data = "data:image/png;base64," + base64.b64encode(f.read()).decode("utf-8")
                return whole_match.replace(url, data) + f" data-hosted-content-id='{hosted_content_id}'"
        else:
            return whole_match

    if msg["body"] and msg["body"]["content"]:
        v = msg["body"]["content"]
        if v:
            if v[0:3].lower() != "<p>":
                v = f"<p>{v}</p>"

            v = re.sub('<emoji.+?alt="(.+?)".+?></emoji>', r"\g<1>", v)

            v = re.sub('<attachment id="(.+?)"></attachment>', get_attachment, v)

            # loosey-goosey matching here :(
            v = re.sub('src="(.+?)"', get_image, v)
        return v

    return None


def render_chat(chat: Dict, output_dir: str):
    """
    render a single chat to an html file. returns the name of the file rendered.
    """

    # read all the msgs for the chat, order them in chron order

    html_dir = os.path.join(output_dir, "html")
    chat_dir = os.path.join(output_dir, "data", chat["id"])

    messages_files = sorted(glob.glob(os.path.join(chat_dir, f"msg_*.json")))
    msgs = []
    for path in messages_files:
        with open(path, "r") as f:
            msg = json.loads(f.read())
            msgs.append({"obj": msg, "content": render_message_body(msg, chat_dir, html_dir)})

    # write out the html file

    filename = f"{chat['id']}.html"

    path = os.path.join(html_dir, filename)

    with open(path, "w") as f:
        print(f"Writing {path}")
        template = get_jinja_env().get_template("chat.jinja")
        f.write(
            template.render(
                chat=chat,
                member_list_str=get_member_list(chat),
                messages=msgs,
            )
        )
    return filename


def render_all(output_dir):
    """render all the chats to html files"""

    all_chats = []

    makedir(os.path.join(output_dir, "html"))

    chat_files = sorted(glob.glob(os.path.join(output_dir, "data", "*.json")))
    for path in chat_files:
        with open(path, "r") as f:
            chat = json.loads(f.read())

            filename = render_chat(chat, output_dir)

            chat_name = get_chat_name(chat)

            all_chats.append({ "filename": filename, "chat_name": chat_name })

    all_chats = sorted(all_chats, key=lambda d: d['chat_name'])

    index_file = os.path.join(output_dir, "html", "index.html")

    with open(index_file, "w") as f:
        print(f"Writing {index_file}")
        template = get_jinja_env().get_template("index.jinja")
        f.write(
            template.render(
                chats=all_chats,
            )
        )


def get_graph_client() -> GraphServiceClient:
    credential = InteractiveBrowserCredential(client_id=pnp_management_shell_client_id)
    scopes = ["Chat.Read"]
    client = GraphServiceClient(credentials=credential, scopes=scopes)
    return client


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=["download", "generate_html"])
    parser.add_argument("--output-dir", type=str, default="archive")
    parser.add_argument("--force", help="download all msgs, not just 'newest' ones", action="store_true")
    args = parser.parse_args()

    if args.command == "download":
        asyncio.run(download_all(args.output_dir, args.force))
    elif args.command == "generate_html":
        render_all(args.output_dir)
    else:
        print(f"Error: unrecognized command '{args.command}'")


if __name__ == "__main__":
    main()
