
import json
from datetime import datetime, timedelta

def init_path():

    with open('config.json', 'r', encoding='utf-8') as f: # открытие файла для чтения
        data = json.load(f) # чтение данных из файла в формате JSON в объект Python

    path = data["Path"]

    return path

def init_path_report():

    with open('config.json', 'r', encoding='utf-8') as f: # открытие файла для чтения
        data = json.load(f) # чтение данных из файла в формате JSON в объект Python

    globalPath = data["Path_report"]

    return globalPath

def init_date():

    with open('config.json', 'r', encoding='utf-8') as f: # открытие файла для чтения
        data = json.load(f) # чтение данных из файла в формате JSON в объект Python

    date = data["Date"]

    return date

def init_otladka():

    with open('config.json', 'r', encoding='utf-8') as f: # открытие файла для чтения
        data = json.load(f) # чтение данных из файла в формате JSON в объект Python

    flag_otladka = data["Otladka"]

    return flag_otladka

def init_email_send():

    with open('config.json', 'r', encoding='utf-8') as f: # открытие файла для чтения
        data = json.load(f) # чтение данных из файла в формате JSON в объект Python
    mas_email=[]
    mas_email.append(data["email1"])
    mas_email.append(data["email2"])
    mas_email.append(data["email3"])

    return mas_email

def update_time(date_message):
    with open('config.json', 'r', encoding='utf-8') as f: # открытие файла для чтения
        data = json.load(f) # чтение данных из файла в формате JSON в объект Python

    dt = datetime.strptime(date_message, '%Y-%m-%d %H:%M:%S')
    dt = dt + timedelta(seconds=5)
    date_message = datetime.strftime(dt, '%Y-%m-%d %H:%M:%S')
    data["Date"] = date_message

    with open('config.json', 'w') as f:
        json.dump(data, f)



#def record_new_date():

mas = []

mas.insert(0, init_path())
mas.insert(1, init_date())
mas.insert(2, init_otladka())
mas.append(init_email_send())

#print(mas)

#print(mas[3][0])
"""path = init()[0]
date = init()[1]
print(path)
print(date)"""