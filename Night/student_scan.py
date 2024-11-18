import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from additional import scope, credentials_file, spreadsheet_id_schedule
def excel_scan_student(day_now, credentials_bool = True):
    if credentials_bool:
        credentials = ServiceAccountCredentials.from_json_keyfile_name('./'+credentials_file, scopes=scope)
    else:
        credentials = ServiceAccountCredentials.from_json_keyfile_name('../'+credentials_file, scopes=scope)
    client = gspread.authorize(credentials)
    sheet = client.open_by_key(spreadsheet_id_schedule)
    file1 = sheet.get_worksheet(0)
    data = file1.get_notes()
    data_name = file1.col_values(col=2)

    training_name = ['Обучение 1', 'Обучение 2', 'Обучение 3', 'Обучение ICL']
    training_excel = ['1', '2', '3', 'основа']
    training = {}
    for i in range(len(training_excel)):
        training[training_excel[i]] = training_name[i]
    result = {}
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
                  'Верле Каролина Валерьевна', 'Севостьянова Софья Евгеньевна', 'Довбышева Анна Павловна',
                  'Левен Елизавета Александровна', 'Каменщиков Александр Александрович', 'Шартнер Карина Аркадьевна']
    col = 1 + 3*(day_now)
    k = 1
    print(col)
    for i in data:
        if len(i) >= col+1 and str(i[col]).find('уз') == 0:
            print(str(i[col]).split('\n')[0])
            if k == 139:
                result[training[str(i[col]).split('\n')[0].split()[2]]] = [k + 2, data_name[k - 2]]
            else:
                result[training[str(i[col]).split('\n')[0].split()[2]]] = [k+1, data_name[k-2]]
        k+=1
    print('student_result: ', result)
    return result