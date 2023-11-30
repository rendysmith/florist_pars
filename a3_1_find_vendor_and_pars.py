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
import a_convert_xls_to_xlsx

warnings.simplefilter("ignore")

abspath = os.path.dirname(os.path.abspath(__file__))
script_path = f'{abspath}/scripts/pdf_parser'

# def new_module(name_script):
#     exec(open(f'{abspath}/pdf_parser/{name_script}', encoding='utf-8').read())
#     read_excel_file(excel_file)
#
# def new_module_2(name_script):
#     exec(open(f'{abspath}/pdf_parser/{name_script}', encoding='utf-8').read())
#     #read_excel_file(excel_file)
#     pdf_convertor_to_xml(file_name, company_name)

def gs_to_json():
    path = f'{abspath}/service_account.json'
    id_sheet = '1PiGYRLX2yXBqqD8SKkyS1_0FKpFpZoGvglog08j6RJM'
    worksheet_name = 'Лист1'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(path)  # Ключ доступа
    gc = gspread.authorize(credentials)  # Открыть соединение
    workbook = gc.open_by_key(id_sheet)  # Открыть таблицу
    worksheet = workbook.worksheet(worksheet_name)  # Открыть лист
    data = worksheet.get_all_values()  # Получить данные из листа
    #data_json = json.dumps(data)  # Преобразовать данные из списка в формат JSON
    #df = pd.DataFrame(data)
    print('Данные из гугла получены')
    return pd.DataFrame(data)

def is_pdf(excel_file):
    abs = os.path.abspath(excel_file)
    company_name = os.path.dirname(abs)
    dir_list = os.listdir(company_name)

    base_name_xlsx = os.path.basename(excel_file)
    name1, ext1 = os.path.splitext(base_name_xlsx)

    for file_d in dir_list:
        name2, ext2 = os.path.splitext(file_d)
        if name1 == name2 and ext2.lower() == '.pdf':
            pdf_file = os.path.join(company_name, file_d)
            return pdf_file

    return False

def write_in_df(df_rep, excel_file):
    folder_path, file_name = os.path.split(excel_file)
    print("Folder path:", folder_path)
    print("File name:", file_name)
    file_name, ext = os.path.splitext(file_name)
    print("File extension:", ext)

    list_files = os.listdir(folder_path)
    for file in list_files:
        file2, ext2 = os.path.splitext(file)
        if file_name == file2 and ext2.lower() == '.pdf':
            len_df = len(df_rep)
            df_rep.loc[len_df, 'File_name'] = file
            break

    #len_df = len(df_rep)
    #df_rep.loc[len_df, 'File_name'] = data

def raspars_to_xml(script_name, excel_file):
    name_file, ext_file = os.path.splitext(excel_file)
    if ext_file.lower() == '.xls':
        if os.path.isfile(name_file + '.xlsx'):
            excel_file = name_file + '.xlsx'
            print('    4 - Файл заменен на новый.')

    print(f'----------------------------------\nExcel_file {excel_file}')
    abspath = os.path.dirname(os.path.abspath(__file__))
    path_script = f'{abspath}/scripts/pdf_parser/{script_name}'
    print('Path_script', path_script)

    spec = importlib.util.spec_from_file_location(script_name, path_script)
    module = importlib.util.module_from_spec(spec)

    send_mail = 0

    try:
        spec.loader.exec_module(module)

        send_mail = 0
        try:
            module.read_excel_file(excel_file)

        except AttributeError as AE:
            print('!!!AE', AE)
            try:
                print(f'     4 - Не найдена основная функция read_excel_file, пробуем подключиться к дополнительной.')

                # file_name_base = os.path.basename(excel_file)
                # pdf_file = os.path.splitext(file_name_base)[0] + '.pdf'

                abs = os.path.abspath(excel_file)
                company_name = os.path.dirname(abs)
                print(f'     4 - Company_name:', company_name)
                dir_list = os.listdir(company_name)
                #print('     4 - Dir_list', dir_list)

                base_name_xlsx = os.path.basename(excel_file)
                name1, ext1 = os.path.splitext(base_name_xlsx)

                for file_d in dir_list:
                    name2, ext2 = os.path.splitext(file_d)
                    #print('     4 - NN2, EE2', name1, name2, ext1, ext2)
                    if name1 == name2 and ext2.lower() == '.pdf':
                        pdf_file = file_d
                        print(f'     4 - Pdf_file найден! -', pdf_file)
                        print(f'     4 - Полный путь до файла: {company_name}/{pdf_file}')
                        module.pdf_convertor_to_excel(pdf_file, company_name)
                        print(f'     4 -------------------')
                        break

            except Exception as Ex:
                print('!!!Ex', Ex)
                subject = 'ERROR! Ошибка в скрипте!'
                print(subject)
                body = f'''<p>
                В скрипте {script_name} обнаружена следующая ошибка<br />
                Строка с ошибкой: {Ex.__traceback__.tb_lineno}<br />
                Данные по ошибке: {Ex}<br />
                Название скрипта: {script_name}<br />
                Название файла: {os.path.basename(excel_file)}<br />
                Решение: Необходимо внести изменение в скрипт для устранения ошибки.<br />
                </p>'''
                send_mail = 1

        except IndexError as IE:
            print('!!!IE', IE)
            subject = 'ERROR! Ошибка в скрипте!'
            print(subject)
            body = f'''<p>
            При обработке файла в скрипте {script_name} произошла ошибка<br />
            Данные по ошибке: {IE}<br />
            Название скрипта: {script_name}<br />
            Название файла: {os.path.basename(excel_file)}<br />
            Решение: Необходимо внести изменение в скрипт файл для устранения ошибки.<br />
            </p>'''
            send_mail = 1

        except Exception as EX:
            print('!!!EX ERROR:', EX.__traceback__.tb_lineno)
            print('!!!EX ERROR:', EX)
            subject = 'ERROR! Данная ошибка пока не распознана!'
            print(subject)
            body = f'''<p>
            В скрипте {script_name} найдена неизвестная пока ошибка<br />
            Строка с ошибкой: {EX.__traceback__.tb_lineno}<br />
            Данные по ошибке: {EX}<br />
            Название скрипта: {script_name}<br />
            Название файла: {os.path.basename(excel_file)}<br />
            Решение: Необходимо внести изменение в скрипт для устранения ошибки.<br />
            </p>'''
            send_mail = 1

        if send_mail == 1:
            count = 0
            while count < 5:
                try:
                    pdf_file = is_pdf(excel_file)
                    if pdf_file != False:
                        a0_modul_mail.send_mail_and_file(subject, body, pdf_file)
                    else:
                        a0_modul_mail.send_mail_and_file(subject, body, excel_file)
                    write_in_df(df_report, excel_file)
                    break

                except:
                    time.sleep(3)
                    count += 1

    #Скрипт отсутствует в папке.
    except FileNotFoundError as FNFE:
        print('FNFE', FNFE)
        subject = 'ERROR! Сервис не смог найти нужный файл со скриптом.'
        print(subject)
        body = f'''<p>
        Сервис не смог обнаружить необходимый скрипт для распознования файла в папке со скриптами<br />
        Данные по ошибке: {FNFE}<br />
        Название скрипта: {script_name}<br />
        Название файла: {os.path.basename(excel_file)}<br />
        Решение: Необходимо поместить скрипт в общую папку со всеми скриптами.<br />
        </p>
        '''
        pdf_file = is_pdf(excel_file)
        if pdf_file != False:
            a0_modul_mail.send_mail_and_file(subject, body, pdf_file)
        else:
            a0_modul_mail.send_mail_and_file(subject, body, excel_file)
        write_in_df(df_report, excel_file)

    except Exception as EXC:
        print('!!!EX ERROR:', EXC)
        subject = 'ERROR! Ошибка при открытии скрипта!'
        print(subject)
        body = f'''<p>
        В скрипте {script_name} найдена ошибка<br />
        Строка с ошибкой: {EXC.__traceback__.tb_lineno}<br />
        Данные по ошибке: {EXC}<br />
        Название скрипта: {script_name}<br />
        Название файла: {os.path.basename(excel_file)}<br />
        Решение: Необходимо внести изменение в скрипт для устранения ошибки.<br />
        </p>'''
        a0_modul_mail.send_mail_and_file(subject, body, excel_file)
        write_in_df(df_report, excel_file)

    #exec(open(path_script, encoding='utf-8').read())

def find_vendor(path_to_xlsx, folder_name):
    print(f'\n4 - Поиск вендора \nПуть до файла: {path_to_xlsx}\nИмя: {folder_name}')
    #print(df)
    clients = df[0].to_list()

    if folder_name != '':
        print(f'4 - Имя вендора {folder_name}')
        try:
            try:
                idx = df.loc[df[0] == folder_name].index[0]
                script_name = df.loc[idx, 1]
                return script_name

            except:
                #return False
                read_file = pd.read_excel(path_to_xlsx, header=None)
                file_rows = read_file.to_string().replace(' ', '').lower()

                for client in clients:
                    client_mod = client.replace(' ', '').lower()
                    if client_mod in file_rows:
                        print(f'    4 - Имя вендора {client}')
                        idx = df.loc[df[0] == client].index[0]
                        script_name = df.loc[idx, 1]
                        return script_name

        except:
            try:
                new_path_to_xlsx = a_convert_xls_to_xlsx.xls_convert_to_xlsx(path_to_xlsx)
                if new_path_to_xlsx != False:
                    read_file = pd.read_excel(new_path_to_xlsx, header=None)
                    file_rows = read_file.to_string().replace(' ', '').lower()

                    for client in clients:
                        client_mod = client.replace(' ', '').lower()
                        if client_mod in file_rows:
                            print(f'    4 - Имя вендора {client}')
                            idx = df.loc[df[0] == client].index[0]
                            script_name = df.loc[idx, 1]
                            return script_name

            except:
                print(' 4 - Файл excel не удалось прочитать.')
                subject = f'ERROR! Файл exсel не смогли прочитать.'
                print(subject)
                body = f'''<p>
                Файл не был прочитан.<br />
                Возможно файл создан в устаревшей программе и/или в устаревшем формате.<br />
                Название файла: {os.path.basename(path_to_xlsx)}<br />
                Решение: Пересохранить файл в новом формате.</p>'''
                a0_modul_mail.send_mail_and_file(subject, body, path_to_xlsx)
                write_in_df(df_report, path_to_xlsx)
                return False

    else:
        try:
            try:
                read_file = pd.read_excel(path_to_xlsx, header=None)
                file_rows = read_file.to_string().replace(' ', '').lower()

                for client in clients:
                    client_mod = client.replace(' ', '').lower()
                    if client_mod in file_rows:
                        print(f'    4 - Имя вендора {client}')
                        idx = df.loc[df[0] == client].index[0]
                        script_name = df.loc[idx, 1]
                        return script_name

            except:
                new_path_to_xlsx = a_convert_xls_to_xlsx.xls_convert_to_xlsx(path_to_xlsx)
                print('    4 - New convertor API')
                if new_path_to_xlsx != False:
                    print('    4 - New convertor API - NO False')
                    read_file = pd.read_excel(new_path_to_xlsx, header=None)
                    file_rows = read_file.to_string().replace(' ', '').lower()

                    for client in clients:
                        client_mod = client.replace(' ', '').lower()
                        if client_mod in file_rows:
                            print(f'    4 - Имя вендора {client}')
                            idx = df.loc[df[0] == client].index[0]
                            script_name = df.loc[idx, 1]
                            return script_name

        except:
            print(' 4 - Файл excel не удалось прочитать.')
            subject = f'ERROR! Файл exсel не смогли прочитать.'
            print(subject)
            body = f'''<p>
            Файл не был прочитан.<br />
            Возможно файл создан в устаревшей программе и/или в устаревшем формате.<br />
            Название файла: {os.path.basename(path_to_xlsx)}<br />
            Решение: Пересохранить файл в новом формате.</p>'''
            a0_modul_mail.send_mail_and_file(subject, body, path_to_xlsx)
            write_in_df(df_report, path_to_xlsx)
            return False

    return False

def find_script(temp_path, file_name, folder_name):
    print('     4 ----- Поиск вендора. -----')
    path_to_xlsx = os.path.join(temp_path, file_name)
    #Поиск имени скрипта в
    script_name = find_vendor(path_to_xlsx, folder_name)
    if script_name != False and script_name != '':
        try:
            print(f'    4 - Найден скрипт в таблице: ->{script_name}<-')
            raspars_to_xml(script_name, path_to_xlsx)
        except ModuleNotFoundError as MNFE:
            subject = f'ERROR! Ошибка функции - {MNFE}.'
            print(subject)
            body = f'''<p>
            Файл не был распознан.<br />
            Возможно не установлена необходимая библиотека для работы скрипта.<br />
            Данные по ошибке: {MNFE}<br />
            Название скрипта: {script_name}<br />
            Название файла: {os.path.basename(path_to_xlsx)}<br />
            Решение: Проверить наличие необходимой библиотеки на сервере, в случаи необходимости доустановить.<br />
             </p>'''
            pdf_file = is_pdf(path_to_xlsx)
            if pdf_file != False:
                a0_modul_mail.send_mail_and_file(subject, body, pdf_file)
            else:
                a0_modul_mail.send_mail_and_file(subject, body, path_to_xlsx)
            write_in_df(df_report, path_to_xlsx)

    else:
        if folder_name != '':
            dop = f'ВНИМАНИЕ! Если инвойс находился в файле eml, необходимо проверить, был ли адрес вендора <{folder_name}> в таблице'
        else:
            dop = ''
        subject = f'ERROR! Файл не распознан.'
        print(subject)
        body = f'''<p>
        Файл не был распознан.<br />
        Возможно не найден вендор в файле или в таблице у вендора не указан скрипт.<br />
        Название файла: {os.path.basename(path_to_xlsx)}<br />
        Решение: Необходимо проверить xlsx файл на наличие данных о вендоре и сопоставить эти данные стаблицей на соответствие.<br />
        {dop}<br />
         </p>'''
        pdf_file = is_pdf(path_to_xlsx)
        if pdf_file != False:
            a0_modul_mail.send_mail_and_file(subject, body, pdf_file)
        else:
            a0_modul_mail.send_mail_and_file(subject, body, path_to_xlsx)
        write_in_df(df_report, path_to_xlsx)

def convert_to_xml(temp_path):
    global df_report
    df_report = pd.DataFrame(columns=['File_name'])
    global df
    df = gs_to_json()

    folders = [e for e in os.listdir(temp_path) if os.path.isdir(os.path.join(temp_path, e))]
    for folder in folders:
        temp_email_path = os.path.join(temp_path, folder)
        files = [i for i in os.listdir(temp_email_path) if os.path.splitext(i)[1].lower() in ['.xlsx', '.xls']]
        for file_name in files:
            find_script(temp_email_path, file_name, folder)

    #Поиск файлов которые не в папках
    files = [i for i in os.listdir(temp_path) if os.path.splitext(i)[1].lower() in ['.xlsx', '.xls']]
    for file_name in files:
        find_script(temp_path, file_name, '')

    # for root, dirs, files in os.walk(temp_path):
    #     for file in files:
    #         print(os.path.join(root, file)

    #if len(df_report) != 0:
    file_name = os.path.basename(os.path.normpath(temp_path))
    file_report = os.path.join(temp_path, f'{file_name}.xlsx')
    df_report.to_excel(file_report, sheet_name='report')
    print('     - Отчет подготовлен')

#convert_to_xml(r'd:\YandexDisk\Python\GitHub\florist\att\20232141013')
