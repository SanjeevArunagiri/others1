from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import os

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def create_service():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)

def get_label_id(service, label_name):
    try:
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        for label in labels:
            if label['name'] == label_name:
                return label['id']
    except Exception as e:
        print(f"An error occurred while fetching label ID: {e}")
        return None
        
def print_labels(service):
    try:
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        for label in labels:
            print(f"Label ID: {label['id']}, Label Name: {label['name']}")
    except Exception as e:
        print(f"An error occurred while fetching labels: {e}")


def main():
    service = create_service()
    print_labels(service)
    label_id = get_label_id(service, 'spam')
    if label_id:
        results = service.users().messages().list(userId='me', labelIds=[label_id]).execute()
        # Process the results as needed
        print(results)
    else:
        print("Label 'fulltime-job-applications' not found.")

if __name__ == '__main__':
    main()
