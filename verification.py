from datetime import date

def verific(time_act, total_time_jira, project_act, project_jira, date_act_f):
    global flag_time
    flag_time = 0
    if time_act == total_time_jira:
        flag_time = 1
        text1 = 'Трудозатраты - Верно'
    else:
        text1 = 'Трудозатраты - Неверно (Проверьте, пожалуйста, трудозатраты, данные в акте должны совпадать с данными указанными в jira за указанный период)'


    global flag_project
    flag_project = 0
    project_act = project_act.replace(' ', '')
    project_act = project_act.lower()
    project_jira = project_jira.replace(' ', '')
    project_jira = project_jira.lower()
    if project_act == project_jira:
        flag_project = 1
        text2 = '- Наименование проекта - Верно'
    else:
        text2 = '- Наименование проекта - Неверно (Проверьте, пожалуйста, наименование проекта, оно должно совпадать в акте и в jira)'

    global flag_date_act
    flag_date_act = 0
    curent_date = date.today()

    if date_act_f == curent_date:
        flag_date_act = 1
        text3 = '- Дата акта - Верно'
    else:
        text3 = '- Дата акта - Неверно (Дата акта должна соответствовать текущей дате)'

    if flag_time == 1 and flag_project == 1 and flag_date_act ==1:
        text_message = 'Завершена роботизированная проверка акта и заявки.\nРезультат обработки: Согласовано.'
    else:
        text_message = text1 + '\n\n' + text2 + '\n\n' + text3

    return text_message