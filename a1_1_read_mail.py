import datetime
import imaplib
import pprint
import threading
import time

import pandas as pd
import requests
from googleapiclient.discovery import build

import os
import pickle
# Gmail API utils
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type
from apiclient import errors
import base64
from base64 import urlsafe_b64decode
from bs4 import BeautifulSoup
import a_get_new_json

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = "dlyainvoce@gmail.com"
user_id = our_email
abspath = os.path.dirname(os.path.abspath(__file__))


def sp_decode(sp):
    if 'text/plain' in sp._headers[0][1]:
        sp = sp.get_payload(decode=True)
        sp = sp.decode('utf8', errors='ignore').replace(u'\u202f', ' ')
        sp = sp.encode('cp1251', 'ignore')
        sp = sp.decode('cp1251')
    else:
        sp = ''
    return sp

def read_body(msg):
    mime_msg = email.message_from_bytes(base64.urlsafe_b64decode(msg['raw']))
    dmsgsubject, dmsgsubjectencoding = email.header.decode_header(mime_msg['Subject'])[0]
    msgsubject = dmsgsubject.decode(*([dmsgsubjectencoding] if dmsgsubjectencoding else [])) if isinstance(
        dmsgsubject, bytes) else dmsgsubject
    print('Subject: ' + msgsubject)
    # Body = 111111111111111111111111111111
    print("----------------------------------------------------")
    body = ""
    if mime_msg.is_multipart():
        for part in mime_msg.walk():
            if part.is_multipart():
                for subpart in part.get_payload():
                    if subpart.is_multipart():
                        for subsubpart in subpart.get_payload():
                            body = body + sp_decode(subsubpart)
                    else:
                        body = body + sp_decode(subpart)
            else:
                body = body + sp_decode(part)
    else:
        body = body + str(mime_msg.get_payload(decode=True)) + '\n'

    print(body)
    return body

def get_mails(temp_folder, minutes):
    unix_min_15 = time.time() - minutes * 60

    def gmail_authenticate():
        creds = None
        # the file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first time
        if os.path.exists(f"{abspath}/token.pickle"):
            with open(f"{abspath}/token.pickle", "rb") as token:
                creds = pickle.load(token)
        # if there are no (valid) credentials availablle, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(f'{abspath}/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # save the credentials for the next run
            with open(f"{abspath}/token.pickle", "wb") as token:
                pickle.dump(creds, token)
        return build('gmail', 'v1', credentials=creds)

    # get the Gmail API service
    service = gmail_authenticate()
    result = service.users().messages().list(userId='me').execute()

    # We can also pass maxResults to get any number of emails. Like this:
    # result = service.users().messages().list(maxResults=200, userId='me').execute()
    messages = result.get('messages')

    def GetAttachments(temp_folder, service, user_id, msg_id, wht_list):  # store_dir):
        """Get and store attachment from Message with given id.
        Args:
          service: Authorized Gmail API service instance.
          user_id: User's email address. The special value "me"
          can be used to indicate the authenticated user.
          msg_id: ID of Message containing attachment.
          store_dir: The directory used to store attachments.
        """
        try:
            message = service.users().messages().get(userId=user_id, id=msg_id).execute()
            # print(message)
            # print(message['payload'])
            date_mail = message['payload']['headers']
            date_mail = [date['value'] for date in date_mail if date['name'] == 'Date'][0]
            print('Date_mail', date_mail)

            if 'GMT' in date_mail:
                given_date = date_mail[:-4]
            elif '+' in date_mail:
                given_date = date_mail[:-6]

            formated_date = datetime.datetime.strptime(given_date, "%a, %d %b %Y %H:%M:%S")
            unix_timestamp = datetime.datetime.timestamp(formated_date)
            #print(unix_timestamp)
            #print(datetime.datetime.fromtimestamp(unix_timestamp).strftime('%Y-%m-%d %H:%M:%S'))

            if unix_timestamp >= unix_min_15:
                #print('Берем в работу')

                att_name = message['payload']['parts'][0]['filename']
                print('Att_name:', att_name)
                #input()

                for part in message['payload']['parts']:
                    print('     Этап 3 - File OLD name', part['filename'])
                    if part['filename']:
                        print('     Этап 3 1')
                        attachment = service.users().messages().attachments().get(userId='me',
                                                                                  messageId=message['id'],
                                                                                  id=part['body']['attachmentId']
                                                                                  ).execute()
                        print('     Этап 3 2\n')  # ', attachment)
                        file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                        print('     Этап 3 3 - Сохраняем файл на диск')
                        mail_file_name = part['filename'].replace('/', '').replace(':', '')

                        if wht_list == 'ok':
                            mail_file_name = mail_file_name.lower().replace('packing', '')
                            if os.path.basename(mail_file_name).lower() == '.eml':
                                mail_file_name = 'WHITE_LIST' + mail_file_name
                                print('     Этап 3 3 1 - Packing file')

                        print('     Этап 3 3 - File NEW name:', mail_file_name)
                        path = f"{temp_folder}/{mail_file_name}"

                        print('     Этап 3 3 Path:', path)
                        if os.path.isfile(path):
                            time.sleep(2)
                            path = f"{temp_folder}/{int(time.time())}_{mail_file_name}"

                        print('     Этап 3 3 New Path:', path)
                        f = open(path, 'wb')
                        f.write(file_data)
                        f.close()

            else:
                print('SKIP MESSANGE')

        except errors.HttpError as error:
            print(f'An error occurred: {error}')

        except Exception as ex:
            print('ERROR!!!\n', ex)
            # pprint.pprint(message['payload'])
            # input()

    white_list = a_get_new_json.get_white_list()

    def read_msg(message):
        print('============================================================================')
        #msg = service.users().messages().get(userId='me', id=message['id']).execute()
        msg = service.users().messages().get(userId=user_id, id=message['id']).execute()
        #msg = service.users().messages().get(userId=user_id, id=message['id'], format='raw').execute()
        subject_headers = [header['value'] for header in msg['payload']['headers'] if header['name'] == 'Subject']
        if len(subject_headers) > 0:
            subject = subject_headers[0]
        else:
            subject = '<no subject>'
        print('Subject:', subject)

        msg_body = service.users().messages().get(userId=user_id, id=message['id'], format='raw').execute()
        decoded_data = read_body(msg_body)

        wht_list = ''
        if any(white in decoded_data for white in white_list):
            wht_list = 'ok'
            print(f'wht_list = {wht_list}')

        else:
            print('*********************************')
            print(white_list)
            print(decoded_data)
            #print('   Pause.....')
            #input('*********************************')

        msg_id = message['id']

        if 'ERROR' not in subject:
            GetAttachments(temp_folder, service, user_id, msg_id, wht_list)

        else:
            try:
                vendor_name = subject.split('*')[1]
                temp_folder2 = f'{temp_folder}/{vendor_name}'
                try:
                    os.mkdir(temp_folder2)
                except:
                    pass

                GetAttachments(temp_folder2, service, user_id, msg_id, wht_list)
            except:
                pass

    threads = []
    for message in messages:
        read_msg(message)
        #input('\nPause...')
    #     thread = threading.Thread(target=read_msg, args=(message, ))
    #     thread.start()
    #     threads.append(thread)
    #     time.sleep(0.5)
    #
    # for thread in threads:
    #     thread.join()
