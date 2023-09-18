# Import `load_workbook` module from `openpyxl`
from openpyxl import load_workbook
from datetime import date
import time
from Parse_word import name_act #передаем данные из акта

def create_report():
    # Load in the workbook
    wb = load_workbook('C:\\Users\\Admin\\PycharmProjects\\AVA\\report.xlsx')

    # Get sheet names
    print(wb.get_sheet_names())

    # Get a sheet by name
    sheet = wb.get_sheet_by_name('Лист1')

    # Print the sheet title
    print(sheet.title)

    # Retrieve the value of a certain cell
    sheet['A2'].value = name_act
    wb.save(f"C:\\Users\\Admin\\PycharmProjects\\AVA\\report\\{date.today()}_{time.time()}_{name_act}.xlsx")

    # Select element 'B2' of your sheet
    sheet['B2'] = 'eee'
















"""import datetime
from datetime import date
from datetime import time
import time
from Parse_word import name_act #передаем данные из акта
import json
import os"""

"""def create_report():
    directory_folder = rf"C:\\Users\\Admin\\PycharmProjects\\AVA\\report\\{date.today()}_{time.time()}_{name_act}.json"
    folder_path = os.path.dirname(directory_folder) # Путь к папке с файлом

    if not os.path.exists(folder_path): #Если пути не существует создаем его
        os.makedirs(folder_path)

    with open(directory_folder, 'w') as file: # Открываем фаил и пишем
        file.write("этот текст создан автоматически")

    data = {}
    data['act'] = []
    data['jira'] = []
    data['act'].append({
        'name': name_act,
        'time': 'pythonist.ru',
        'raid': 'Nebraska'
    })
    data['jira'].append({
        'name': 'Larry',
        'time': 'pythonist.ru'
    })

    with open(f'C:\\Users\\Admin\\PycharmProjects\\AVA\\report\\file_{date.today()}_{time.time()}_{name_act}.json', 'w', "utf-8") as outfile:
        json.dump(data, outfile)"""

