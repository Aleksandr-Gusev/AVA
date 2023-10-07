import imaplib
import email
from email.header import decode_header
import base64
import re


mail = imaplib.IMAP4_SSL('mail.flexcloud.ru')
mail.login('actbot@i-sol.ru', 'Parol1!')

mail.list()
mail.select("inbox")

#result, data = mail.search(None, "SINCE 04-Oct-2023 SUBJECT 'AVA'")
#result, response = mail.search(None, 'SINCE 04-Oct-2023', '(SUBJECT "AVA")', '(UNSEEN)')
result, msg = mail.search(None, '(SUBJECT "AVA")', '(UNSEEN)')
unread_msg_nums = msg[0].split()
print(unread_msg_nums)
print(len(unread_msg_nums))


id_list = []
for e_id in unread_msg_nums:
    msg = mail.fetch(e_id, '(UID BODY[TEXT])')
    id_list.append(msg[0][1])
    raw_email = msg[0][1]
    raw_email_string = raw_email.encode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    print(email.utils.parseaddr(email_message['From']))
print(id_list)

#print(email_message.utils.parseaddr(email_message['From']))

for e_id in unread_msg_nums:
    mail.store(e_id, '+FLAGS', '\Seen')




"""result, response = mail.fetch(None, '(SUBJECT "AVA")', '(UNSEEN)')
ids = response[0]
id_list = ids.split()
latest_email_id = id_list[-1]
raw_email = response[0][1]
raw_email_string = raw_email.decode('utf-8')"""