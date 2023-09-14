import docx
import os

# -----------------------------------------------поиск и запись всех путей файлов------------------------
paths = []
folder = os.path.dirname('C:\\Users\\Admin\\PycharmProjects\\AVA\\Acts')
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('docx') and not file.startswith('~'):
            paths.append(os.path.join(root, file))
print('Всего файлов', len(paths))
print(paths)

# -----------------------------------------------создание экземпляра документа------------------------
for path in paths:
    doc = docx.Document(path)
    properties = doc.core_properties
    """print('Автор документа:', properties.author)
    print('Автор последней правки:', properties.last_modified_by)
    print('Дата создания документа:', properties.created)
    print('Дата последней правки:', properties.modified)
    print('Дата последней печати:', properties.last_printed)
    print('Количество сохранений:', properties.revision)"""

# -----------------------------------------------работа с таблицами------------------------
    mas_tables=[]
    for table in doc.tables:
        mas_tables.append(table)
        """for row in table.rows:
            for cell in row.cells:
                print(cell.text)"""

    print(len(mas_tables))

# -----------------------------------------------присвоение переменных------------------------
    number_act = 0
    number_zayavka = 0
    date_act = ''
    project_act = ''
    key_project_act = ''
    period = ''
    time_act = 0
    rate_act = 0
    rate_zayavka = 0
    project_cost_act = 0
    total_cost_act = 0
    total_cost_zayavka = 0

    number_act = mas_tables[0].cell(0, 0).text[-1]  # Номер акта
    number_zayavka = mas_tables[4].cell(0, 0).text[-1] # Номер заявки
    date_act = mas_tables[1].cell(0, 1).text  # Дата акта
    project_act = mas_tables[2].cell(1, 1).text  # Наименование проекта
    key_project_act = mas_tables[2].cell(1, 2).text  # Ключ проекта
    period = mas_tables[2].cell(1, 4).text  # период
    time_act = mas_tables[2].cell(1, 5).text  # трудозатраты
    rate_act = mas_tables[2].cell(1, 6).text  # ставка в акте
    rate_zayavka = mas_tables[5].cell(1, 1).text  # ставка в заявке
    project_cost_act = mas_tables[2].cell(1, 7).text  # стоимость по проекту
    total_cost_act = mas_tables[2].cell(2, 7).text  # Итого в акте
    total_cost_zayavka = mas_tables[5].cell(1, 3).text  # Итого в заявке
    name_act2 = mas_tables[3].cell(1, 1).text[3:len(mas_tables[3].cell(1, 1).text)]  # имя в  акте 2
    name_act3 = mas_tables[6].cell(1, 1).text[3:len(mas_tables[3].cell(1, 1).text)]  # имя  в акте 3


    print(number_act)
    print(number_zayavka)
    print(date_act)  # Дата акта
    print(project_act)  # Наименование проекта
    print(key_project_act)  # Ключ проекта
    print(period)  # период
    print(time_act)  # трудозатраты
    print(rate_act)  # ставка в Акте
    print(rate_zayavka)  # ставка в заявке
    print(project_cost_act)  # стоимость по проекту
    print(total_cost_act)  # Итого
    print(total_cost_zayavka)
    print(name_act2)
    print(name_act3)

# -----------------------------------------------работа с текстом------------------------
text = []
for paragraph in doc.paragraphs:                   #получение списка параграфов
    text.append(paragraph.text)

"""print('\n'.join(text))"""



# -----------------------------------------------выделение имени из акта------------------------
full_name_act1 = text[5][text[5].index('ИП')+3:text[5].index(', именуемый')] #выделение имени
print(full_name_act1)

name_act = full_name_act1[0:full_name_act1.rfind(' ')]
name_act = name_act.replace(' ', '')
print(name_act)

# -----------------------------------------------выделение суммы из текста акта------------------------


rub = text[9][text[9].index('сумму ')+6:text[9].index(' (')] #выделение суммы руб
copeyka = text[9][text[9].index(' копе')-2:text[9].index(' копе')] #выделение суммы копейки
print(rub)
print(copeyka)

total_cost_act_in_text = rub + ',' + copeyka # полная сумма
print(total_cost_act_in_text)
