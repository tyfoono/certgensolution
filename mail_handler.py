import base64
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/gmail.send']
MESSAGE_TEMPLATES = {'en': 'Hello, %s! Thanks for your participation. The attachment includes your %s certification.',
                     'ru': "Здравствуйте, %s! Спасибо за участие. В прикрепленном файле документ, подтверждающий результат вашего участие в %s"}


def send_gmail(email: str, name: str, language: str, event_name: str, path: str = "") -> None:
    if not os.path.exists('credentials.json'):
        print("Error: credentials.json file not found")
        return None

    creds = None

    if os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        except Exception as e:
            print(f"Error loading token: {e}")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Refresh failed: {e}")
                # If refresh fails, delete the token file and start over
                if os.path.exists('token.json'):
                    os.remove('token.json')
                creds = None

        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                print(f"Authentication failed: {e}")
                return None

        # Save credentials only if we have valid ones
        try:
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        except Exception as e:
            print(f"Error saving token: {e}")
            return None

    try:
        service = build('gmail', 'v1', credentials=creds)

        # Create a multipart message for potential attachments
        message = MIMEMultipart()
        message['to'] = email
        message['from'] = 'satunogerm@gmail.com'
        message['subject'] = event_name

        # Add the text body
        body = MIMEText(MESSAGE_TEMPLATES[language] % (name, event_name))
        message.attach(body)

        # Add attachment if path is provided and file exists
        if os.path.exists(path):
            if os.path.exists(path):
                with open(path, 'rb') as file:
                    attachment = MIMEApplication(file.read(), _subtype='pdf')
                    attachment.add_header('Content-Disposition', 'attachment',
                                          filename=path.split('/')[-1])
                    message.attach(attachment)

        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        result = service.users().messages().send(
            userId='me',
            body={'raw': raw_message}
        ).execute()

        print(f"Message sent to {email}. Message ID: {result['id']}")

    except HttpError as error:
        print(f"HTTP Error: {error}")
        print(f"Status code: {error.resp.status}")
        print(f"Error details: {error.error_details}")
    except Exception as error:
        print(f"Unexpected error: {error}")
