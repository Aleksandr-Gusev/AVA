import imaplib
import email
from email.header import decode_header
import base64
import re
import os
from imbox import Imbox # pip install imbox
import traceback
from datetime import datetime
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

    date_str = time.strftime("%Y-%m-%d %H:%M:%S", date_msg)
    #date_today = datetime.today().strftime("%Y-%m-%d %H:%M:%S")
    #date_today = "2023-10-17 16:00:00"
    date_today = init_date()            # получение даты из конфигурационного файла

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

