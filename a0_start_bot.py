import datetime
import os, sys
import shutil
import time
import requests

import a0_modul_mail
import a1_1_read_mail
import a1_2_unzip_eml
import a2_1_convert_pdf
import a3_1_find_vendor_and_pars
import a3_2_send_report
import a4_1_send_file_to_gdisk

def main(minutes):
    txt_mail = ''
    txt_disk = ''

    now = datetime.datetime.now()
    temp_folder = now.strftime("%Y%m%d_%H%M")

    abspath = os.path.dirname(os.path.abspath(__file__))
    print('Abspath', abspath)
    try:
        os.mkdir(f'{abspath}/att')
    except:
        pass

    temp_path = f'{abspath}/att/{temp_folder}'
    try:
        os.mkdir(temp_path)
    except:
        pass

    a0_modul_mail.send_telegram(f'+++ 🟢Start! {time.ctime()}, period = {minutes}')
    print("Temp_path", temp_path)
    txt = f"1 - Скачивание писем с почты каждые {minutes} мин, историю писем за {minutes} мин."
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    time.sleep(3)
    nm = 0
    while nm < 5:
        try:
            a1_1_read_mail.get_mails(temp_path, minutes) #Проверка каждый 15 мин.
            break
        except Exception as EX:
            a0_modul_mail.send_telegram(f"1 - ERROR\n{EX.__traceback__.tb_lineno}\n{EX}")
            nm += 1
            time.sleep(10)
            #sys.exit()
        #
    #input('Пока стоп')
    txt = "2 - вытаскиваем данные из eml файлов"
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    try:
        a1_2_unzip_eml.unzip_and_save(temp_path)
    except Exception as EX:
        a0_modul_mail.send_telegram(f"2 - ERROR {EX}")
        sys.exit()

    txt = "3 - опрос новых папок для распознования файлов в excel файлы."
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    try:
        a2_1_convert_pdf.converter(temp_path)
    except Exception as EX:
        a0_modul_mail.send_telegram(f"3 - ERROR {EX}")
        sys.exit()

    txt = "4 - Поиск вендора по файлу"
    #a3_1_find_vendor_and_pars.convert_to_xml(temp_path)
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    n = 0
    while n <= 3:
        try:
            a3_1_find_vendor_and_pars.convert_to_xml(temp_path)
            break

        except Exception as EX:
            a0_modul_mail.send_telegram(f"4 - ERROR\n{EX.__traceback__.tb_lineno}\n{EX}")
            n += 1
            time.sleep(10)
            #sys.exit()

    txt = "5 - Копирование файлов формата xml на google disk"
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    try:
        a4_1_send_file_to_gdisk.find_xml_files(temp_path)
        txt_disk = '\n- - - - Файлы на диск загружены.'
    except Exception as EX:
        a0_modul_mail.send_telegram(f"6 - ERROR {EX}")
        sys.exit()

    txt = "6 - Подготовка отчета по работе"
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    try:
        a3_2_send_report.find_for_report(temp_path)
        txt_mail = '\n- - - - Отчет отправлен.'
    except Exception as EX:
        a0_modul_mail.send_telegram(f"5 - ERROR {EX}")
        sys.exit()

    txt = "7 - Удалить временную папку со всем содержимым"
    print(txt)
    #a0_modul_mail.send_telegram(txt)
    #shutil.rmtree(temp_path)
    if not os.listdir(temp_path):
        # Если папка пустая, то удалите ее
        os.rmdir(temp_path)
        print("Папка удалена успешно!")
    else:
        shutil.rmtree(temp_path)
        print("Папка не является пустой. Всё равно удалить)")

    a0_modul_mail.send_telegram(f'- - - 🔴Stop! {time.ctime()}{txt_mail}{txt_disk}')

now = datetime.datetime.now()

# current_time = time.localtime()
#
# print(now, current_time)

print(now.hour, now.minute)

if 9 <= now.hour <= 11:
    if now.minute < 14:
        if now.hour == 9:
            main(16)

        elif now.hour in [10, 11]:
            main(31)

    elif now.hour == 11 and now.minute >= 15:
        main(16)

    elif 30 <= now.minute <= 44:
        main(31)

elif 18 <= now.hour <= 22:
    if now.minute < 14:
        if now.hour == 18:
            main(16)

        elif now.hour in [19, 20, 21, 22]:
            main(31)

    elif now.hour == 22 and now.minute >= 15:
        main(16)

    elif 30 <= now.minute <= 44:
        main(31)

elif 4 <= now.hour <= 5:
    if now.hour == 4 and now.minute < 14:
        main(16)

    elif now.hour == 5 and now.minute < 14:
        main(61)

    elif now.hour == 5 and now.minute >= 15:
        main(16)

else:
    main(16)


