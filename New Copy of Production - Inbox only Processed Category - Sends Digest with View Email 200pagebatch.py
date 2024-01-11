from O365 import Account, FileSystemTokenBackend, Message
from datetime import datetime, timedelta
import pytz
import html
import requests
import json

# Configuration details
client_id = 'TBD'
client_secret = 'TBD'
user_email = 'derek.jaindl@tka.com'
resource_email = 'TBD@tka.com'
tenant_id = 'tbd'
token_file_path = 'tbd'
Team_email1 = 'as@tka.com'
Team_email2 = 'ac@tka.com'
Team_email3 = 're@tka.com'

# Setting token backend
token_backend = FileSystemTokenBackend(token_path='C:\\Users\\djaindl-c\\ChatBot', token_filename='o365_token_updated.txt')

# Authenticate account
credentials = (client_id, client_secret)
account = Account(credentials, token_backend=token_backend, tenant_id=tenant_id)

processed_category = "Processed"  # Name of the category to mark processed emails

try:
    if not account.is_authenticated:
        account.authenticate(scopes=['https://graph.microsoft.com/Mail.Read', 
                                     'https://graph.microsoft.com/Mail.Send',
                                     "profile",
                                     "openid",
                                     "email",
                                     "https://graph.microsoft.com/Mail.Read.Shared",
                                     "https://graph.microsoft.com/Mail.ReadWrite",
                                     "https://graph.microsoft.com/Mail.ReadWrite.Shared",
                                     "https://graph.microsoft.com/User.Read",
                                     "https://graph.microsoft.com/Mail.Send.Shared"])

    # Access mailbox and folders
    mailbox = account.mailbox(resource=resource_email)
    inbox = mailbox.inbox_folder()

    page_size = 200  # Number of emails to fetch per page

    while True:
        # Refresh the token before each batch
        account.connection.refresh_token()
        access_token = account.connection.token_backend.token['access_token']
        headers = {'Authorization': 'Bearer ' + access_token}

        # Fetch all emails directly from the Inbox
        messages = inbox.get_messages(limit=page_size)
        message_list = list(messages)  # Convert generator to list

        if not message_list:
            print("No new emails to process. Exiting.")
            break

        html_content = "<h1>Email Summary</h1>"

        for message in message_list:
            # Skip messages that have already been processed
            if processed_category in message.categories:
                continue

            sender = message.sender.address
            received = message.received.strftime('%Y-%m-%d %H:%M:%S')
            subject = html.escape(message.subject)
            preview = html.escape(message.body_preview)
            link = html.escape(message.web_link)

            # Append each message's details to the HTML content
            html_content += f"""
                <div style="margin-bottom: 20px;">
                    <strong>From:</strong> {sender}<br>
                    <strong>Received:</strong> {received}<br>
                    <strong>Subject:</strong> {subject}<br>
                    <p>{preview}</p>
                    <p><a href="{link}">View Email</a></p>
                </div>
            """

            # Update the categories of the message
            update_payload = {'categories': message.categories + [processed_category]}
            update_url = f'https://graph.microsoft.com/v1.0/users/{resource_email}/messages/{message.object_id}'
            response = requests.patch(update_url, headers=headers, json=update_payload)

            if response.status_code not in [200, 204]:
                print(f"Failed to update message: {response.json()}")

        if html_content.strip() == "<h1>Email Summary</h1>":
            print("No new emails to include in the summary. Exiting.")
            break

        # Prepare and send an email with the HTML content for each batch
        new_message = account.new_message()
        new_message.to.add(user_email)
        new_message.to.add(Team_email1)
        new_message.to.add(Team_email2)
        new_message.to.add(Team_email3)
        new_message.subject = "Email Summary"
        new_message.body = html_content
        new_message.body_type = 'HTML'
        new_message.send()

        print("Summary email sent for a batch.")

        # Check if less than page_size emails were processed, indicating the end
        if len(message_list) < page_size:
            print("Processed all available emails. Exiting.")
            break

except Exception as e:
    print(f"An error occurred: {e}")