# server.py — Outlook MCP Server

from mcp.server.fastmcp import FastMCP
import msal
import requests

from dotenv import load_dotenv
import os

load_dotenv()
CLIENT_ID = os.getenv("CLIENT_ID")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Mail.ReadWrite", "Mail.ReadWrite.Shared", "Mail.Send"]
TOKEN_CACHE_FILE = "token_cache.json"

mcp = FastMCP("Outlook Mail")


def get_access_token() -> str:
    """Gets a valid access token using saved login."""
    cache = msal.SerializableTokenCache()
    try:
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    except FileNotFoundError:
        raise Exception("Please run auth.py first!")

    app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache
    )
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0])

    with open(TOKEN_CACHE_FILE, "w") as f:
        f.write(cache.serialize())

    return result["access_token"]


@mcp.tool()
def get_unread_emails() -> str:
    """Fetches the latest 10 unread emails from Outlook inbox. Returns sender, subject and date."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        "https://graph.microsoft.com/v1.0/me/messages"
        "?$filter=isRead eq false"
        "&$select=subject,from,receivedDateTime,bodyPreview"
        "&$orderby=receivedDateTime desc"
        "&$top=10"
    )
    data = requests.get(url, headers=headers).json()
    emails = data.get("value", [])

    if not emails:
        return "You have no unread emails! 🎉"

    result = f"You have {len(emails)} unread email(s):\n\n"
    for i, email in enumerate(emails, 1):
        sender = email["from"]["emailAddress"]["address"]
        subject = email["subject"]
        date = email["receivedDateTime"][:10]
        preview = email["bodyPreview"][:100]
        result += f"{i}. From: {sender}\n   Subject: {subject}\n   Date: {date}\n   Preview: {preview}...\n\n"
    return result


@mcp.tool()
def get_unread_count() -> str:
    """Returns the total number of unread emails in the Outlook inbox."""
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "ConsistencyLevel": "eventual"
    }
    url = "https://graph.microsoft.com/v1.0/me/messages/$count?$filter=isRead eq false"
    count = requests.get(url, headers=headers).text
    return f"You have {count} unread email(s) in your Outlook inbox."


@mcp.tool()
def send_email(to: str, subject: str, body: str) -> str:
    """Sends an email from your Outlook account. to=recipient email, subject=email subject, body=email content."""
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    email_data = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to}}]
        }
    }
    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers=headers,
        json=email_data
    )
    if response.status_code == 202:
        return f"✅ Email sent successfully to {to}!"
    return f"❌ Failed to send email. Error: {response.text}"


@mcp.tool()
def move_email_to_folder(email_subject: str, folder_name: str) -> str:
    """Moves an email to a specified folder. email_subject=subject to search for, folder_name=destination folder name."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Find the email
    search_url = f"https://graph.microsoft.com/v1.0/me/messages?$filter=contains(subject,'{email_subject}')&$top=1"
    email_data = requests.get(search_url, headers=headers).json()
    emails = email_data.get("value", [])

    if not emails:
        return f"❌ No email found with subject containing '{email_subject}'"

    email_id = emails[0]["id"]

    # Find the folder
    folders_data = requests.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders",
        headers=headers
    ).json()
    folders = folders_data.get("value", [])
    folder = next((f for f in folders if f["displayName"].lower() == folder_name.lower()), None)

    if not folder:
        return f"❌ Folder '{folder_name}' not found. Available folders: {', '.join(f['displayName'] for f in folders)}"

    # Move the email
    move_url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/move"
    response = requests.post(
        move_url,
        headers={**headers, "Content-Type": "application/json"},
        json={"destinationId": folder["id"]}
    )

    if response.status_code == 201:
        return f"✅ Email moved to '{folder_name}' successfully!"
    return f"❌ Failed to move email. Error: {response.text}"


@mcp.tool()
def flag_email_for_followup(email_subject: str) -> str:
    """Flags an email for follow-up based on subject. email_subject=subject to search for."""
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Find the email
    search_url = f"https://graph.microsoft.com/v1.0/me/messages?$filter=contains(subject,'{email_subject}')&$top=1"
    email_data = requests.get(search_url, headers=headers).json()
    emails = email_data.get("value", [])

    if not emails:
        return f"❌ No email found with subject containing '{email_subject}'"

    email_id = emails[0]["id"]

    # Flag it
    flag_url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
    response = requests.patch(
        flag_url,
        headers=headers,
        json={"flag": {"flagStatus": "flagged"}}
    )

    if response.status_code == 200:
        return f"✅ Email flagged for follow-up!"
    return f"❌ Failed to flag email. Error: {response.text}"

@mcp.tool()
def mark_all_emails_as_read() -> str:
    """Marks all unread emails in the Outlook inbox as read."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Get all unread emails
    url = (
        "https://graph.microsoft.com/v1.0/me/messages"
        "?$filter=isRead eq false&$select=id&$top=50"
    )
    emails = requests.get(url, headers=headers).json().get("value", [])

    if not emails:
        return "No unread emails to mark!"

    # Mark each one as read
    count = 0
    for email in emails:
        response = requests.patch(
            f"https://graph.microsoft.com/v1.0/me/messages/{email['id']}",
            headers={**headers, "Content-Type": "application/json"},
            json={"isRead": True}
        )
        if response.status_code == 200:
            count += 1

    return f"✅ Marked {count} emails as read!"


# ─── RESOURCES ────────────────────────────────────

@mcp.resource("emails://inbox/summary")
def inbox_summary() -> str:
    """A summary of your current Outlook inbox status."""
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "ConsistencyLevel": "eventual"
    }

    # Get unread count
    count = requests.get(
        "https://graph.microsoft.com/v1.0/me/messages/$count?$filter=isRead eq false",
        headers=headers
    ).text

    # Get latest 3 senders
    emails = requests.get(
        "https://graph.microsoft.com/v1.0/me/messages"
        "?$filter=isRead eq false&$select=from,subject&$top=3",
        headers={"Authorization": f"Bearer {token}"}
    ).json().get("value", [])

    senders = [e["from"]["emailAddress"]["address"] for e in emails]

    return f"""
Inbox Summary:
- Unread emails: {count}
- Latest senders: {', '.join(senders) if senders else 'None'}
- Last checked: just now
"""


@mcp.resource("emails://inbox/folders")
def inbox_folders() -> str:
    """Lists all available folders in your Outlook inbox."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    folders = requests.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders",
        headers=headers
    ).json().get("value", [])

    result = "Your Outlook Folders:\n\n"
    for f in folders:
        result += f"- {f['displayName']} ({f['unreadItemCount']} unread)\n"
    return result


@mcp.resource("emails://inbox/flagged")
def flagged_emails() -> str:
    """Lists all emails you have flagged for follow-up."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    emails = requests.get(
        "https://graph.microsoft.com/v1.0/me/messages"
        "?$filter=flag/flagStatus eq 'flagged'"
        "&$select=subject,from,receivedDateTime&$top=10",
        headers=headers
    ).json().get("value", [])

    if not emails:
        return "No flagged emails! You're all caught up. ✅"

    result = f"You have {len(emails)} flagged email(s):\n\n"
    for i, e in enumerate(emails, 1):
        result += f"{i}. {e['subject']} — from {e['from']['emailAddress']['address']}\n"
    return result




@mcp.prompt()
def daily_inbox_review() -> str:
    """Morning inbox review — summarizes unread emails and suggests actions."""
    return """
    Please do a complete morning inbox review for me:
    
    1. First call get_unread_count to see how many unread emails I have
    2. Then call get_unread_emails to fetch them
    3. Group the emails by type:
       - 🔴 URGENT (anything about security, payments, deadlines, interviews)
       - 📋 ACTION NEEDED (things that need a reply or decision)
       - 📰 INFO ONLY (newsletters, notifications, digests)
    4. For URGENT and ACTION NEEDED emails, suggest what to do
    5. End with a 1-line summary like: "You have X urgent, Y need action, Z are info-only."
    
    Be concise and clear. Format with emojis for easy scanning.
    """


@mcp.prompt()
def email_reply_helper(sender: str, subject: str) -> str:
    """Helps draft a professional reply to a specific email."""
    return f"""
    I need to reply to an email:
    - From: {sender}
    - Subject: {subject}
    
    Please help me:
    1. Suggest what this email is likely about based on the subject/sender
    2. Draft 2 different reply options:
       - Option A: Short and direct (3-4 sentences)
       - Option B: Detailed and professional (full paragraph)
    3. Ask me which one I want to send, then use send_email to send it
    
    Wait for my choice before sending anything.
    """



if __name__ == "__main__":
    mcp.run()

