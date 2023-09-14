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

    date_act = mas_tables[1].cell(0, 1).text  # Дата акта
    project_act = mas_tables[2].cell(1, 1).text  # Наименование проекта
    key_project_act = mas_tables[2].cell(1, 2).text  # Ключ проекта
    period = mas_tables[2].cell(1, 4).text  # период
    time_act = mas_tables[2].cell(1, 5).text  # трудозатраты
    rate_act = mas_tables[2].cell(1, 6).text  # ставка
    project_cost_act = mas_tables[2].cell(1, 7).text  # стоимость по проекту
    total_cost_act = mas_tables[2].cell(2, 7).text  # Итого

    print(date_act)  # Дата акта
    print(project_act)  # Наименование проекта
    print(key_project_act)  # Ключ проекта
    print(period)  # период
    print(time_act)  # трудозатраты
    print(rate_act)  # ставка
    print(project_cost_act)  # стоимость по проекту
    print(total_cost_act)  # Итого

# -----------------------------------------------работа с текстом------------------------
"""text = []
for paragraph in doc.paragraphs:
    text.append(paragraph.text)
print('\n'.join(text))"""