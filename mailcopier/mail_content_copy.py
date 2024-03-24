from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
from bs4 import BeautifulSoup
from docx import Document

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API to fetch inbox
    results = service.users().messages().list(userId='me',labelIds=['Label_xyz']).execute()
    messages = results.get('messages', [])

    if not messages:
        print('No new mails.')
    else:
        message_count = 0
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            email_data = msg['payload']['headers']
            for values in email_data:
                name = values['name']
                if name == 'From':
                    from_name = values['value']
                    for part in msg['payload']['parts']:
                        try:
                            data_text = part['data']
                            data = base64.urlsafe_b64decode(data_text)
                            soup = BeautifulSoup(data , "lxml")
                            body = soup.body()
                            # Create a Document
                            doc = Document()
                            doc.add_paragraph(body)
                            doc.add_page_break()
                            doc.save('mails.docx')
                            message_count += 1
                        except Base64DecodeError:
                            pass
        print(f'{message_count} mails saved.')
if __name__ == '__main__':
    main()