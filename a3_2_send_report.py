import pandas as pd
import os, sys
import time
import warnings
import gspread
from pprint import pprint
from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
import importlib.util
import a0_modul_mail

warnings.simplefilter("ignore")

abspath = os.path.dirname(os.path.abspath(__file__))
script_path = f'{abspath}/scripts/pdf_parser'

def find_for_report(temp_path):
    pdf_num = 0
    xlsx_num = 0
    xls_num = 0
    xml_num = 0

    for root, dirs, files in os.walk(temp_path):
        for filename in files:
            file_ext = os.path.splitext(filename)[1].lower()
            if file_ext == '.pdf':
                pdf_num += 1
            elif file_ext == '.xml':
                xml_num += 1
            elif file_ext == '.xlsx':
                xlsx_num += 1
            elif file_ext == '.xls':
                xls_num += 1

    if pdf_num != 0:
        file_name = os.path.basename(os.path.normpath(temp_path))
        file_report = os.path.join(temp_path, f'{file_name}.xlsx')

        df = pd.read_excel(file_report)
        error_num = len(df)

        subject = f'Report: Отчет от {time.ctime()}'
        body = f'''<p>
Отчет от {time.ctime()}<br />
Входящих файлов pdf = {pdf_num}<br />
Файлов xlsx = {xlsx_num}<br />
Файлов xls = {xls_num}<br />
Файлов распознано и сконвертировано = {xml_num}<br />
Файлов с ошибкой = {error_num}<br />
Список нераспознаных документов приложен в файле.<br />
</p>'''

        print(body)
        print(file_report)
        a0_modul_mail.send_mail_and_file(subject, body, file_report)
        print('     - Отчет отправлен.')

#find_for_report(r'd:\YandexDisk\Python\GitHub\florist\att\20230311_0800')
