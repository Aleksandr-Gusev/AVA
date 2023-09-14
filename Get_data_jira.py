from __future__ import annotations
import re
from collections import Counter
from typing import cast
from jira import JIRA
from jira.client import ResultList
from jira.resources import Issue
import requests
from datetime import datetime


jira = JIRA(server='https://jira.i-sol.eu', basic_auth=('tcontrol', '1J*iBJGT'))


#------------------------ Ввод ФИО---------------------
date1 = '01.08.2023'
date2 = '31.08.2023'

date1 = date1.replace('.', '/')
date2 = date2.replace('.', '/')
date1 = datetime.strptime(date1, "%d/%m/%Y")
date2 = datetime.strptime(date2, "%d/%m/%Y")
print(date1)
print(date2)
#name_user = input('Введите Фамилию и Имя: \n')
name_user = "Девятайкина Анна"
name_user = name_user.replace(" ", "")
#name_project = input('Введите ключ проекта: \n')
name_project = 'EVRAZTMC'

#----------------------- получение задач проекта-----------------------
issues_in_proj = jira.search_issues(f'project={name_project}')
proj = jira.project('EVRAZTMC')

print(proj.name) # имя проекта
print(issues_in_proj)

#----------------------- поиск задач в которых есть ФИО-----------------------

for j in range(len(issues_in_proj)):
    x = jira.worklogs(issue=issues_in_proj[j])

    time = 0

    print(issues_in_proj[j])
    for i in range(len(x)):
        author = x[i].updateAuthor.displayName
        str = author.replace(" ", "")
        index = x[i].started.split('T')  # разделение даты от времени
        date_jira = datetime.strptime(index[0], "%Y-%m-%d")  # дата джиры

        if str == name_user and date1<=date_jira and date2>=date_jira:
            print(x[i].started)
            print(index)
            print(x[i].id)
            #print(x[i].timeSpentSeconds)
            #print(x[i].updateAuthor)
            time = x[i].timeSpentSeconds + time
    print(time/3600)




