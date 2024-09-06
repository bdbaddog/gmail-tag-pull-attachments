import os
import base64
import re
import io
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email import message_from_bytes

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def authenticate_gmail():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


def search_messages_by_label(service, labelId):
    try:
        # Call the Gmail API to search for messages with the specified query
        results = service.users().messages().list(userId='me', labelIds=labelId, maxResults=500).execute()
        messages = results.get('messages', [])
        return messages
    except HttpError as error:
        print(f'An error occurred: {error}')
        return []

def get_label_id(service, label):
    try:
        results = service.users().labels().list(userId='me').execute()
        for labelDict in results['labels']:
            if labelDict['name'] == label:
                return labelDict['id']
    except HttpError as error:
        print(f'An error occurred: {error}')
        return []
def download_attachments(service, message_id, output_dir):
    try:
        # Get the message using the Gmail API
        message = service.users().messages().get(userId='me', id=message_id).execute()

        # Iterate through the parts of the message
        parts = message['payload'].get('parts', [])
        for part in parts:
            filename = part.get('filename')
            if filename:
                # Check if the part has an attachment
                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    att_id = part['body'].get('attachmentId')
                    if att_id:
                        att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                           id=att_id).execute()
                        data = att['data']
                    else:
                        continue
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                path = os.path.join(output_dir, filename)

                with open(path, 'wb') as f:
                    f.write(file_data)
                print(f"Downloaded: {filename}")

        if not parts:
            att_id = message['payload']['body'].get('attachmentId', False)
            filename = message['payload'].get('filename', False)
            if att_id:
                att = service.users().messages().attachments().get(userId='me', messageId=message_id,id=att_id).execute()
                data = att['data']
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                path = os.path.join(output_dir, filename)

                with open(path, 'wb') as f:
                    f.write(file_data)
                print(f"Downloaded: {filename}")
            else:
                print("something wrong")

    except HttpError as error:
        print(f'An error occurred: {error}')


def main():
    # Authenticate and build the Gmail service
    service = authenticate_gmail()

    # Specify your DMARC filter query
    label = 'dmarc'

    # Specify the directory to save attachments
    output_dir = './attachments'

    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    label = get_label_id(service, 'DMARC/ParcTel')
    print(f"Found ID: {label}")
    # Search for messages containing "dmarc" in the subject or body
    messages = search_messages_by_label(service, label)

    if not messages:
        print("No messages found.")
    else:
        for message in messages:
            download_attachments(service, message['id'], output_dir)


if __name__ == '__main__':
    main()

