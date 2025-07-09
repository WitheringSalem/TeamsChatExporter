
import requests
import msal
import os
from datetime import datetime
from pathlib import Path
import hashlib
import re


# Requires a 365 App Registration with the permissions 'AuditLog.Read.All', 'Chat.Read.All' & 'User.Read.All'
# Replace with your Azure AD app credentials
CLIENT_ID = '<>'
TENANT_ID = '<>'
CLIENT_SECRET = '<>'
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPES = ['https://graph.microsoft.com/.default']
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

users_for_export = [
    '<>']

# Setup MSAL app
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)



# Acquire token
token_response = app.acquire_token_for_client(scopes=SCOPES)
access_token = token_response['access_token']
headers = {
'Authorization': f'Bearer {access_token}'
}

# Output folder
output_dir = Path('teams_chat_export')
output_dir.mkdir(exist_ok=True)

def get_users():
    url = f'{GRAPH_ENDPOINT}/users'
    all_users = []

    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            print(f"Failed to get users: {response.status_code} - {response.text}")
            break

        data = response.json()
        all_users.extend(data.get('value', []))
        url = data.get('@odata.nextLink')  # Grab next page if it exists

    print(f'Total users retrieved: {len(all_users)}')
    return all_users

def get_user_chats(user_id):
    url = f'{GRAPH_ENDPOINT}/users/{user_id}/chats'
    response = requests.get(url, headers=headers)
    return response.json().get('value', [])

def get_chat_messages(chat_id):
    url = f'{GRAPH_ENDPOINT}/chats/{chat_id}/messages'
    response = requests.get(url, headers=headers)
    return response.json().get('value', [])

def get_chat_members(chat_id):
    url = f'{GRAPH_ENDPOINT}/chats/{chat_id}/members'
    response = requests.get(url, headers=headers)
    return response.json().get('value', [])

def export_chat_to_html(user_name, readable_name, chat_type, chat_id, messages):
    safe_name = readable_name if readable_name else chat_id
    safe_name = safe_name.replace("@", "_at_").replace(".", "_").replace(" ", "_")

    #Truncate if too long (e.g. limit to 100 chars)
    if len(safe_name) > 100:
        # Append a hash to preserve uniqueness
        digest = hashlib.sha256(safe_name.encode('utf-8')).hexdigest()[:8]
        safe_name = f"{safe_name[:90]}__{digest}"

    html = f'<html><head><meta charset="utf-8"><title>{chat_type} - {safe_name}</title></head><body>'
    html += f'<h2>{chat_type.title()} Chat - {safe_name}</h2>'

    for msg in messages:
        from_user = 'Unknown'
        msg_from = msg.get('from')
        if msg_from:
            if 'user' in msg_from:
                from_user = msg_from['user'].get('displayName', 'Unknown User')
            elif 'application' in msg_from:
                from_user = msg_from['application'].get('displayName', 'Unknown App')
            elif 'device' in msg_from:
                from_user = msg_from['device'].get('displayName', 'Unknown Device')
            else:
                from_user = msg_from.get('displayName', 'Unknown Sender')

        timestamp = msg.get('createdDateTime', '')
        content = msg.get('body', {}).get('content', '')

        html += f'<p><strong>{from_user}</strong> [{timestamp}]:<br>{content}</p><hr>'

    html += '</body></html>'

    file_path = output_dir / f'{user_name.replace(" ", "_")}__{safe_name}.html'
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'Exported: {file_path}')

def clean_filename_part(name):
    if not name or not isinstance(name, str):
        name = "Unknown"
    name = name.replace(" ", "_")
    return re.sub(r'[^A-Za-z0-9_]', '', name)


# Main script
def main():
    print('Getting users list')
    users = get_users()

    for user in users:
        user_id = user['id']
        user_principal_name = user['userPrincipalName']

        if user_principal_name in users_for_export:
            user_name = user['displayName']
            print(f'Exporting chats for: {user_name}')
            chats = get_user_chats(user_id)

            for chat in chats:
                chat_id = chat['id']
                chat_type = chat.get('chatType', 'unknown')
                members = get_chat_members(chat_id)

                # Attempt to get a readable title
                member_names = [clean_filename_part(m.get('displayName')) for m in members]
                readable_name = "_".join(member_names)

                messages = get_chat_messages(chat_id)
                export_chat_to_html(user_name, readable_name, chat_type, chat_id, messages)

if __name__ == '__main__':
    main()