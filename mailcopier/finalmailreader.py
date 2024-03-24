from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from bs4 import BeautifulSoup
from docx import Document
from base64 import urlsafe_b64decode
import os
import pickle

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('/Users/sanjiv/Downloads/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

    service = build('gmail', 'v1', credentials=creds)

    try:
        results = service.users().messages().list(userId='me', labelIds=['Label_7854785072520162857']).execute()
        messages = results.get('messages', [])

        if not messages:
            print('No new mails.')
        else:
            message_count = 0
            doc = Document()  # Create a Document

            for i, message in enumerate(messages):
                try:
                    msg = service.users().messages().get(userId='me', id=message['id']).execute()

                    for part in msg['payload']['parts']:
                        try:
                            data_text = part['body']['data']
                            data = urlsafe_b64decode(data_text)
                            soup = BeautifulSoup(data, 'html.parser')
                            body = soup.get_text()

                            doc.add_paragraph(body)
                            if i != len(messages) - 1:  # Don't add a page break if it's the last email
                                doc.add_page_break()

                            message_count += 1
                        except KeyError:
                            print("Error: 'data' key not found in part")
                        except Exception as e:
                            print(f'Error processing message: {e}')

            doc.save('mails.docx')  # Save the Document
            print(f'{message_count} mails saved.')
    except Exception as e:
        print(f'Error calling Gmail API: {e}')

if __name__ == '__main__':
    main()
