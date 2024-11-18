import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from additional import scope, credentials_file, spreadsheet_id_schedule
def excel_scan_wiretapping(coef_name, day_now):
    print('coef_name: ', coef_name)
    print('day_now: ', day_now)
    credentials = ServiceAccountCredentials.from_json_keyfile_name('./'+credentials_file, scopes=scope)
    client = gspread.authorize(credentials)
    sheet = client.open_by_key(spreadsheet_id_schedule)
    file1 = sheet.get_worksheet(0)
    data = file1.get_all_values()
    name_opers = ['Куликов Максим Геннадьевич', 'Плетнёв Владислав Евгеньевич', 'Абдулина Дарья Игоревна',
                 'Баландин Дмитрий Алексеевич', 'Идрисалиева Ольга Ивановна', 'Третьяков Даниел Павлович',
                 'Гарипова Лейсан Миневелиевна', 'Архангельская Елена Николаевна', 'Стаценко Юлия Валентиновна',
                 'Швачка Мария Юрьевна ', 'Гарифуллин Булат Дамирович', 'Яценко Александра Андреевна',
                 'Севостьянов Михаил Игоревич', 'Баринова София Андреевна', 'Шептунова Софья Денисовна',
                 'Джабраилова Елена Валерьевна', 'Миргазова Ляйсан Ришатовна', 'Довыдович Алиса Станиславовна',
                 'Маркелов Сергей Алексеевич', 'Горных Ксения Сергеевна', 'Исаева Екатерина Аркадьевна',
                 'Смертина София Сергеевна', 'Гребенюк Полина Максимовна', 'Шартнер Елена Александровна',
                 'Матылицкая Елена Юрьевна', 'Магомедова Патимат Каримуллаевна ', 'Щербакова Лилана Евгеньевна',
                 'Орищенко Юлия Геннадьевна', 'Серова Юлия Владимировна', 'Алескарова Кристина Фёдоровна',
                 'Шаламова Дарья Сергеевна  ', 'Путко Екатерина Николаевна', 'Нагая Анна Марьевна',
                 'Рыжкина Екатерина Михайловна ', 'Степечев Сергей Сергеевич', 'Бахтиаров Глеб Русланович',
                 'Жданов Денис Александрович', 'Апанасенко Екатерина Григорьевна', 'Апанасенко София Дмитриевна',
                 'Дыдолева Виктория Сергеевна', 'Батталова Алсу Альбертовна', 'Силин Егор Владимирович',
                 'Шестаков Дмитрий Петрович', 'Клочкова София Антоновна', 'Шамли Алина Алексеевна',
                 'Зиберт Ольга Николаевна', 'Сонин Владимир Викторович', 'Мунасипова Надия Мидхатовна',
                 'Штанкевич Ольга Петровна', 'Ханнанова Елизавета Рустамовна ', 'Крушина Наталья Анатольевна',
                 'Верле Каролина Валерьевна']
    col = 5 + 3*(day_now - 1)
    formula_name = {}
    name_cef_in_formul = {}
    name = list(coef_name)
    days = {}
    row = 32
    opers_formul = {}
    for i in range(31, 191):
        if data[i][1] in name and (data[i][3 + 3*(day_now - 1)] or data[i+1][3 + 3*(day_now - 1)]):
            if data[i+3][1] in name_opers:
                opers_formul[data[i][1]] = i+3
            elif data[i+4][1] in name_opers:
                opers_formul[data[i][1]] = i + 4
    print(opers_formul)
    for i in opers_formul:
        formula = str(file1.cell(row=opers_formul[i], col=col, value_render_option='FORMULA').value)
        formula_name[i] = formula
    print(formula_name)
    for i in formula_name:
        if formula_name[i]!='None':
            if formula_name[i][-10] == '+' or formula_name[i][-10] == '-':
                name_cef_in_formul[i] = formula_name[i][:-10]
            elif formula_name[i][-9] == '+' or formula_name[i][-9] == '-':
                name_cef_in_formul[i] = formula_name[i][:-9]
            elif formula_name[i][-8] == '+' or formula_name[i][-8] == '-':
                name_cef_in_formul[i] = formula_name[i][:-8]
            elif formula_name[i][-7] == '+' or formula_name[i][-7] == '-':
                name_cef_in_formul[i] = formula_name[i][:-7]
            elif formula_name[i][-6] == '+' or formula_name[i][-6] == '-':
                name_cef_in_formul[i] = formula_name[i][:-6]
            elif formula_name[i][-5] == '+' or formula_name[i][-5] == '-':
                name_cef_in_formul[i] = formula_name[i][:-5]
            elif formula_name[i][-4] == '+' or formula_name[i][-4] == '-':
                name_cef_in_formul[i] = formula_name[i][:-4]
    final_formul = {}
    # Строка: формула
    for i in name_cef_in_formul:
        final_formul[opers_formul[i]] = name_cef_in_formul[i] + coef_name[i]+';"")'
    print(final_formul)
    # Вставляем формулу
    # for i in final_formul:
    #     file1.update_cell(row=i, col=col, value=final_formul[i])
    return name_cef_in_formul
