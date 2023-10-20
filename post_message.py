
import smtplib                                      # Импортируем библиотеку по работе с SMTP
import os
# Добавляем необходимые подклассы - MIME-типы
from email.mime.multipart import MIMEMultipart      # Многокомпонентный объект
import mimetypes                                          # Импорт класса для обработки неизвестных MIME-типов, базирующихся на расширении файла
from email import encoders                                # Импортируем энкодер
from email.mime.base import MIMEBase                      # Общий тип
from email.mime.text import MIMEText                      # Текст/HTML
from email.mime.image import MIMEImage                    # Изображения
from email.mime.audio import MIMEAudio                    # Аудио
from read_mail_in_date import object_senders
from init import init_path, init_otladka

def get_sender(path, obj):                      #Получение адресса почты кому отправить письмо
    str = init_path()
    name_doc = path.replace(str, '')
    sender = ''
    for key, value in obj.items():
        if key == name_doc:
            sender = value
    return sender



def send_message(text, sub, path, name_act, number_act, date_act, total_cost, flag_OK, type_of_act, project_act):

    sender_email = get_sender(path, object_senders)  # почта отправителя
    addr_from = "actbot@i-sol.ru"                       # Адресат

    if init_otladka() == '1':
        addr_to = "aleksandr.gusev@i-sol.ru" + ',' + 'gusev.rpa@gmail.com'                 # Получатель
        #addr_to = "actbot@i-sol.ru"                             # Получатель

    """else:
        addr_to = "aleksandr.gusev@i-sol.ru" """                       # Получатель
        # addr_to = "actbot@i-sol.ru"                             # Получатель

    password  = "Parol1!"                                  # Пароль

    msg = MIMEMultipart()                               # Создаем сообщение
    msg['From'] = addr_from                          # Адресат
    msg['To'] = addr_to                            # Получатель

    if flag_OK == 1:
        msg['Subject'] = 'Акт согласован - ' + name_act + ' (' + project_act + ')'     # Тема сообщения
    else:
        msg['Subject'] = 'Результат роботизированной проверки акта'                    # Тема сообщения
    body = text
    msg.attach(MIMEText(body, 'plain'))                 # Добавляем в сообщение текст

    filepath = path  # Имя файла в абсолютном или относительном формате
    filename = os.path.basename(filepath)  # Только имя файла


    #print(get_sender(path, object_senders))


    ctype, encoding = mimetypes.guess_type(filepath)            # Определяем тип файла на основе его расширения
    if ctype is None or encoding is not None:                   # Если тип файла не определяется
        ctype = 'application/octet-stream'                      # Будем использовать общий тип
    maintype, subtype = ctype.split('/', 1)                     # Получаем тип и подтип
    if maintype == 'text':                                      # Если текстовый файл
        with open(filepath) as fp:                              # Открываем файл для чтения
            file = MIMEText(fp.read(), _subtype=subtype)         # Используем тип MIMEText
            fp.close()                                          # После использования файл обязательно нужно закрыть
    else:                                                       # Неизвестный тип файла
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)                  # Используем общий MIME-тип
            file.set_payload(fp.read())                         # Добавляем содержимое общего типа (полезную нагрузку)
            fp.close()
            encoders.encode_base64(file)                        # Содержимое должно кодироваться как Base64
    if flag_OK == 1:
        if type_of_act == 0:
            file.add_header('Content-Disposition', 'attachment', filename=f'ИП {name_act} - Акт № {number_act} от {date_act}г. ({total_cost}).docx')  # Добавляем заголовки
        if type_of_act == 1:
            file.add_header('Content-Disposition', 'attachment', filename=f'{name_act} - Акт № {number_act} от {date_act}г. ({total_cost}).docx')  # Добавляем заголовки
    else:
        file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
    msg.attach(file)                                            # Присоединяем файл к сообщению

    server = smtplib.SMTP('mail.flexcloud.ru', 587)           # Создаем объект SMTP
    server.set_debuglevel(True)                             # Включаем режим отладки - если отчет не нужен, строку можно закомментировать
    server.starttls()                                       # Начинаем шифрованный обмен по TLS
    server.login(addr_from, password)                       # Получаем доступ
    server.send_message(msg)                                # Отправляем сообщение
    server.quit()                                           # Выходим