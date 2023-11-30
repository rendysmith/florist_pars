import datetime
import json
import eml_parser
import pprint
import os
import email
from email import policy
from email.parser import BytesParser
from eml_parser import EmlParser
import warnings
import chardet

warnings.simplefilter("ignore")

def unzip_eml(temp_folder, name_eml):
    eml_path = os.path.join(temp_folder, name_eml)

    with open(eml_path, 'rb') as f:
        result = chardet.detect(f.read())
        encoding = result['encoding']
        print(f'     2.1 - Encoding = {encoding}')

    with open(eml_path, 'r', encoding=encoding) as f:
        msg = email.message_from_file(f)
        from_email = msg['From']
        print('     2.1 - From_email:', from_email)
        idx1 = from_email.rfind('<')
        idx2 = from_email.rfind('>')
        print('     2.1 - Position < and > in email:', idx1, idx2)
        if idx1 != -1:
            from_email = from_email[idx1+1:idx2]
        print('     2.1 - From:', from_email)

    new_temp_folder = os.path.join(temp_folder, from_email)
    try:
        os.mkdir(new_temp_folder)
    except:
        pass

    with open(eml_path, 'rb') as fhdl:
        msg = email.message_from_binary_file(fhdl)

    for part in msg.walk():
        #print('Part', part)

        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        file_name = part.get_filename()
        print('     2.1 - File OLD name in eml 1:', file_name)
        if bool(file_name):
            file_name = file_name.replace('/', '')
            print('     2.1 - File NEW name in eml 2:', file_name)
            file_path = os.path.join(new_temp_folder, file_name)
            print('     2.1 - File path:', file_path)
            with open(file_path, 'wb') as f:
                f.write(part.get_payload(decode=True))

def unzip_eml_2(temp_folder, name_eml):
    eml_file_path = os.path.join(temp_folder, name_eml)
    with open(eml_file_path, 'rb') as eml_file:
        msg = BytesParser(policy=policy.default).parse(eml_file)
        from_email = msg['From']
        idx1 = from_email.rfind('<')
        idx2 = from_email.rfind('>')
        print('     2.2 - Position < and > in email:', idx1, idx2)
        if idx1 != -1:
            from_email = from_email[idx1+1:idx2]
        print('     2.2 - From email:', from_email)

        new_temp_folder = os.path.join(temp_folder, from_email)
        try:
            os.mkdir(new_temp_folder)
        except:
            pass

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            filename = part.get_filename()
            if not filename:
                continue

            payload = part.get_payload(decode=True)
            if payload is None:
                continue

            filename = filename.replace('/', '')
            if 'WHITE_LIST' in name_eml:
                filename = filename.lower().replace('packing', '')

            print('     2.2 - File NEW name in eml 2:', filename)
            file_path = os.path.join(new_temp_folder, filename)
            print('     2.2 - File path:', file_path)

            with open(file_path, 'wb') as f:
                f.write(payload)


def unzip_and_save(temp_path):
    list_files = os.listdir(temp_path)
    list_files = [i for i in list_files if os.path.splitext(i)[1] == '.eml']
    print(list_files)
    for name_file in list_files:
        print('\n-------------------------------\nName_file eml:',name_file)
        try:
            unzip_eml_2(temp_path, name_file)
            print('---------------1111111111111111-----------------')
        except:
            unzip_eml(temp_path, name_file)    #     unzip_eml(temp_folder, name_file)
            print('----------------22222222222222-------------------')

#unzip_and_save(r'd:\YandexDisk\Python\GitHub\florist\tst')