from datetime import date

def verific(time_act, total_time_jira, project_act, project_jira, date_act_f, full_name_act1, name_act2, name_act3, number_act, number_zayavka, rate_act, rate_zayavka, cost_for_verification, project_cost_act, total_cost_act, total_cost_act_in_text, total_cost_zayavka, date1, date2):
    global flag_time
    global flag_OK
    flag_OK = 0             #флаг если все проверки успешны примет 1
    flag_time = 0
    if time_act == total_time_jira:
        flag_time = 1
        #text1 = 'Трудозатраты - Верно \n\n'
        text1 = ''
    else:
        text1 = '   Трудозатраты - Неверно (проверьте, пожалуйста, трудозатраты, данные в акте должны совпадать с данными в JIRA за указанный период. В JIRA заведено - ' + total_time_jira + ' ч.) \n\n'


    global flag_project
    flag_project = 0
    project_act = project_act.replace(' ', '')
    project_act = project_act.lower()
    project_jira_f = project_jira.replace(' ', '')
    project_jira_f = project_jira_f.lower()
    if project_act == project_jira_f:
        flag_project = 1
        #text2 = 'Наименование проекта - Верно \n\n'
        text2 = ''
    else:
        text2 = '   Наименование проекта - Неверно (проверьте, пожалуйста, наименование проекта. Наименование проекта в JIRA: ' + project_jira + ') \n\n'

    global flag_date_act
    flag_date_act = 0
    curent_date = date.today()

    if date_act_f != -1:
        if date_act_f == curent_date:
            flag_date_act = 1
            #text3 = 'Дата акта - Верно \n\n'
            text3 = ''
        else:
            text3 = '   Дата акта - Неверно (дата акта должна соответствовать текущей дате) \n\n'
    else:
        text3 = '   Дата акта - Неверно (Проверьте, пожалуйста, корректно ли указан месяц в дате акта) \n\n'

#---------------- проверка совпадения ФИО----------------------------------------------------
    global flag_FIO
    flag_FIO = 0
    full_name_act1 = full_name_act1.replace(' ', '')
    full_name_act1 = full_name_act1.lower()
    name_act2 = name_act2.replace(' ', '')
    name_act2 = name_act2.lower()
    name_act3 = name_act3.replace(' ', '')
    name_act3 = name_act3.lower()
    if full_name_act1 == name_act2 and full_name_act1 == name_act3:
        flag_FIO = 1
        text4 = ''
    else:
        text4 = '   Заполнение полей с ФИО - Неверно (проверьте, пожалуйста, корректность заполнения полей с ФИО в акте и заявке) \n\n'
# ---------------- проверка совпадения номера акта и заявки----------------------------------------------------
    global flag_number_act
    flag_number_act = 0

    if number_act == number_zayavka:
        flag_number_act = 1
        text5 = ''
    else:
        text5 = '   Номер акта или заявки - Неверно (проверьте, пожалуйста, номера акта и заявки, они должны быть одинаковыми) \n\n'
# ---------------- проверка совпадения номера акта и заявки----------------------------------------------------
    global flag_rate
    flag_rate = 0
    rate_zayavka = rate_zayavka.replace(' ', '')
    rate_act = rate_act.replace(' ', '')
    if rate_zayavka == rate_act:
        flag_rate = 1
        text6 = ''
    else:
        text6 = '   Ставка в акте или заявке - Неверно (проверьте, пожалуйста, ставки в акте и заявке, они должны быть одинаковыми) \n\n'
# ---------------- проверка совпадения стоимости по проекту----------------------------------------------------
    global flag_cost_project
    flag_cost_project = 0
    project_cost_act = format(float(project_cost_act.replace(',', '.')), '.2f')
    if cost_for_verification == project_cost_act:
        flag_cost_project = 1
        text7 = ''
    else:
        text7 = '   Стоимость проекта - Неверно (проверьте, пожалуйста, расчет стоимости проекта) \n\n'

    # ---------------- проверка совпадения стоимости по Итого в таблице----------------------------------------------------
    global flag_cost_total_table
    flag_cost_total_table = 0
    total_cost_act = format(float(total_cost_act.replace(',', '.')), '.2f')
    if cost_for_verification == total_cost_act:
        flag_cost_total_table = 1
        text8 = ''
    else:
        text8 = '   Итоговая стоимость проекта - Неверно (проверьте, пожалуйста, расчет итоговой стоимости проекта) \n\n'

    # ---------------- проверка совпадения стоимости Итого в тексте----------------------------------------------------
    global flag_cost_total_text
    flag_cost_total_text = 0
    total_cost_act_in_text = format(float(total_cost_act_in_text.replace(',', '.')), '.2f')
    if cost_for_verification == total_cost_act_in_text:
        flag_cost_total_text = 1
        text9 = ''
    else:
        text9 = '   Итоговая стоимость проекта в тексте - Неверно (проверьте, пожалуйста, указанную итоговую стоимость в тексте) \n\n'

    # ---------------- проверка совпадения стоимости в акте и в заявке----------------------------------------------------
    global flag_cost_total_zayvka
    flag_cost_total_zayvka = 0
    if total_cost_zayavka != '-':
        total_cost_zayavka = format(float(total_cost_zayavka.replace(',', '.')), '.2f')
        if cost_for_verification == total_cost_zayavka:
            flag_cost_total_zayvka = 1
            text10 = ''
        else:
            text10 = '   Итоговая стоимость проекта в заявке - Неверно (проверьте, пожалуйста, итоговую стоимость в заявке) \n\n'
    else:
        text10 = ''

    # ---------------- проверка периода проверки----------------------------------------------------
    global flag_period
    flag_period = 0

    if date1 < date2:
        flag_period = 1
        text11 = ''
    else:
        text11 = '   Период оказания услуг - Неверно (проверьте, пожалуйста, период оказания услуг в акте) \n\n'
        text1 = ''

#------------------------формирование текста сообщения------------------------------------------
    if flag_time == 1 and flag_project == 1 and flag_date_act == 1 and flag_FIO == 1 and flag_number_act == 1 and flag_rate == 1 and flag_cost_project == 1 and flag_cost_total_table == 1 and flag_cost_total_text == 1 and flag_period == 1:
        text_message = 'Завершена роботизированная проверка акта и заявки.\nРезультат обработки: Согласовано.'
        flag_OK = 1
    else:
        text_message = 'Завершена роботизированная проверка акта и заявки.\nРезультат обработки:'+ '\n\n' + text1 + text2 + text3 + text4 + text5 + text6 + text7 + text8 + text9 + text10 + text11 +'Акт и заявка во вложении. Просьба проверить документы в соответствии с замечаниями и направить повторно на проверку на электронный адрес actbot@i-sol.ru'

    return [text_message, flag_OK]