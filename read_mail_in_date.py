import imaplib
import email
from email.header import decode_header
import base64
import re
import os
from imbox import Imbox # pip install imbox
import traceback
from email.utils import parsedate_tz, mktime_tz, formatdate
from datetime import datetime, timedelta
import time
from tqdm import tqdm   # установить
from init import init_date



mail_pass = "Parol1!"
username = "actbot@i-sol.ru"
imap_server = "mail.flexcloud.ru"
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(username, mail_pass)

imap.select("INBOX")

res, msg = imap.search(None, '(SUBJECT "AVA")')

#print(msg)
unread_msg_nums = msg[0].split()

#print(unread_msg_nums)
print('>>> Проверка электронной почты и выгрузка актов...')
#print(len(unread_msg_nums))

#res, msg = imap.fetch(unread_msg_nums[0], '(RFC822)')  #Для метода search по порядковому номеру письма
count = 1

object_senders = {}         # словарь - имя файла:адрес почты
for e_id in tqdm(unread_msg_nums):
    #tqdm(count_progress)
    time.sleep(0.01)
    res, msg = imap.fetch(e_id, '(RFC822)')  #Для метода search по порядковому номеру письма

    #letter_date = email.utils.parsedate_tz(msg["Date"]) # дата получения, приходит в виде строк
    msg = email.message_from_bytes(msg[0][1])#, _class = email.message.EmailMessage)

    #print(decode_header(msg['Subject'][0].decode()))
    #print(msg['Subject'])
    #print(msg['X-Mailing-List'])         #тело письма
    #print(msg['Date'])
    #print(msg["Return-path"])
    #print(email.utils.parseaddr(msg['From'])[1])
    #print(email.utils.parsedate(msg['Date']))
    date_msg = email.utils.parsedate(msg['Date'])

    # --------------------------- расчет времени в текущем часовом поясе (костыль, нужно сделать по нормальному)-------------------------------
    tzone = parsedate_tz(msg['Date'])                   #пополучаем форма с часовым поясом в формате 10800 сек.
    tzone_delta = round((tzone[9] - 3*3600)/3600)       #сравнивается текущая зона с зоной в кортеже (10800 это +3000) (вынести текущий часовой пояс в конфиг)
    t_fact = date_msg[3] - tzone_delta                  #вычисляем разницу поясов
    d_fact = date_msg[2]                                #записываем день из кортежа
    t_fact_str = ''
    d_fact_str = ''

    if t_fact>=0 and t_fact<=9:
        if t_fact == 0: t_fact_str = "00"
        if t_fact == 1: t_fact_str = "01"
        if t_fact == 2: t_fact_str = "02"
        if t_fact == 3: t_fact_str = "03"
        if t_fact == 4: t_fact_str = "04"
        if t_fact == 5: t_fact_str = "05"
        if t_fact == 6: t_fact_str = "06"
        if t_fact == 7: t_fact_str = "07"
        if t_fact == 8: t_fact_str = "08"
        if t_fact == 9: t_fact_str = "09"
    if t_fact>9: t_fact_str = str(t_fact)
    if t_fact<0:
        t_fact = 24 + t_fact
        t_fact_str = str(t_fact)
        d_fact_str = str(d_fact - 1)

    #date_msg = email.utils.parsedate_to_datetime(msg['Date'])
    date_str = time.strftime("%Y-%m-%d %H:%M:%S", date_msg)
    #------------------------- если есть различие в часовых поясах----------------------------
    if tzone_delta != 0:                                                                    # если разница часовых зон не равно 0
        date_str = time.strftime(f"%Y-%m-%d {t_fact_str}:%M:%S", date_msg)
        if d_fact_str != '':                                                                # если письмо было прислано фактически было прислано вчера требуется уменьшить день
            date_str = time.strftime(f"%Y-%m-{d_fact_str} {t_fact_str}:%M:%S", date_msg)
    #--------------------------------------------------------------------------------------------
    date_str2 = time.strftime("%Y-%m-%d", date_msg)                             # для типизирования
    date_for_second_verific = datetime.strptime(date_str2, "%Y-%m-%d").date()  # переменная для проверки даты акта и даты прихода письма на почту
    #date_today = datetime.today().strftime("%Y-%m-%d %H:%M:%S")
    #date_today = "2023-10-17 16:00:00"
    date_today = init_date()            # получение даты из конфигурационного файла

    print(msg['Date'])  # здесь зарыта собака
    print(date_today)

    #print(datetime.today())

    if date_str >= date_today:

    #------------------- загрузка вложений ----------------------------------

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if decode_header(filename)[0][1] is not None:
                filename = decode_header(filename)[0][0].decode(decode_header(filename)[0][1])  #декодирование из base64 в utf-8
            if not filename:
                filename = 'filename'
            # ------------------- запись только с вложение .docx ----------------------------------
            if filename.find('.docx') != -1:
                object_senders[f'{count}_{filename}'] = email.utils.parseaddr(msg['From'])[1]
            # -----------------------------------------------------
            save_path = os.path.join('C:\\Users\\Admin\\PycharmProjects\\AVA\\Acts\\', f'{count}_{filename}')
            #save_path = os.path.join('C:\\Users\\Admin\\PycharmProjects\\AVA\\Acts\\', f'{filename}_{count}.docx')
            with open(save_path, 'wb') as f:
                f.write(part.get_payload(decode=True))

        #imap.store(e_id, '+FLAGS', '\Seen')
        count +=1                               # защита от совпадения имени

#print(object_senders)
#print('OK')
#import Parse_word                               # запуск модуля

