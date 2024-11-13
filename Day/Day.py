import datetime
import math
from calendar import monthrange

import pandas as pd

def timetable_day(day_plus = 1):
    spreadsheet_id = '1MQ0hNCotH-ou1jp3ODHeBhdww1w6UhG8X588-n-db6c'
    url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet_id + '/export?format=xlsx&rtpof=true&sd=true'
    operators = ['Куликов Максим Геннадьевич', 'Плетнёв Владислав Евгеньевич', 'Абдулина Дарья Игоревна',
                 'Баландин Дмитрий Алексеевич', 'Идрисалиева Ольга Ивановна', 'Третьяков Даниел Павлович',
                 'Гарипова Лейсан Миневелиевна', 'Архангельская Елена Николаевна', 'Стаценко Юлия Валентиновна',
                 'Швачка Мария Юрьевна ', 'Гарифуллин Булат Дамирович', 'Яценко Александра Андреевна',
                 'Севостьянов Михаил Игоревич', 'Баринова София Андреевна', 'Шептунова Софья Денисовна',
                 'Джабраилова Елена Валерьевна', 'Миргазова Ляйсан Ришатовна', 'Довыдович Алиса Станиславовна',
                 'Маркелов Сергей Алексеевич', 'Горных Ксения Сергеевна', 'Исаева Екатерина Аркадьевна',
                 'Смертина София Сергеевна', 'Гребенюк Полина Максимовна', 'Шартнер Елена Александровна',
                 'Матылицкая Елена Юрьевна', 'Магомедова Патимат Каримуллаевна ', 'Щербакова Лилана Евгеньевна',
                 'Орищенко Юлия Геннадьевна', 'Серова Юлия Владимировна', 'Алескарова Кристина Фёдоровна',
                 'Шаламова Дарья Сергеевна  ', 'Путко Екатерина Николаевна ', 'Нагая Анна Марьевна',
                 'Рыжкина Екатерина Михайловна ', 'Степечев Сергей Сергеевич', 'Бахтиаров Глеб Русланович',
                 'Жданов Денис Александрович', 'Апанасенко Екатерина Григорьевна', 'Апанасенко София Дмитриевна',
                 'Дыдолева Виктория Сергеевна', 'Батталова Алсу Альбертовна', 'Силин Егор Владимирович',
                 'Шестаков Дмитрий Петрович', 'Клочкова София Антоновна ', 'Шамли Алина Алексеевна',
                 'Зиберт Ольга Николаевна', 'Сонин Владимир Викторович', 'Мунасипова Надия Мидхатовна',
                 'Штанкевич Ольга Петровна', 'Ханнанова Елизавета Рустамовна ', 'Крушина Наталья Анатольевна',
                 'Верле Каролина Валерьевна']
    tags = ['@Walking_suicide', '@apatiya444', '@DariaIg0revna', '@dimanite07', '@Qlga20', '@tredap', '@leisanaGAT',
            '@elenaarkhangel', '@yuvsta', '@maria177177177', '@chillxr', '@sssnde7', '@Super_Smi', '@sonssonny',
            '@sonyashark8', '@Liseno4k1', '@LimmLlmm', 'Довыдович$', '@kiksonkik', '@klopiksa', '@kaisa87',
            '@Sonysmertina', '@polina_selfdevelopment', '@Blinn20', 'Elena', '@lissana_23', '@lilanka00',
            '@JuliaOrischenko', '@Marry98765', '@Kristi060', '@D000000000000000000000', '@katin_tg', '@antonkacore',
            '@Ryzhkina_Katya', 'Степечев$', '@The5_55', '@deniszhdvokzal', 'Апанасенко$', 'Апанасенко$', '@Bukaaa_000',
            'Батталова$', '@TKT_TESH', 'Шестаков$', 'Клочкова$', '@aliiiiiiina12', '@Olga_Zibert', '@svv223', '@Надия',
            'Штанкевич$', '@Barashek81', 'Крушина$', '@aivazovskaya8']
    operators_tags = {}
    for i in range(0, len(operators)):
        operators_tags[operators[i]] = tags[i]
    night = ['Куликов Максим Геннадьевич', 'Плетнёв Владислав Евгеньевич', 'Абдулина Дарья Игоревна',
             'Баландин Дмитрий Алексеевич']

    def scan_excel(sheet_name, now_month):
        now_year = datetime.datetime.now().year
        # now_month = datetime.datetime.now().month
        days = monthrange(now_year, now_month)[1]

        df = pd.read_excel(url, sheet_name=sheet_name)
        dw = df.values
        name = []
        date_data = {}
        for i in range(1, days + 1):
            date_data[i] = []
        day = 1
        # for i in range(1, ()len(dw[0])):
        for i in range(1, ((days + 1) * 3)):
            oper_work = {}

            # С ночниками
            if i == 1:
                for j in range(len(dw)):
                    if j < 137 and j > 28 and (j == 29 or (j - 29) % 3 == 0):
                        name.append(dw[j][i])
                    elif j < 184 and j > 137 and (j == 138 or (j - 138) % 3 == 0):
                        name.append(dw[j][i])
            elif (i - 3) % 3 == 0:
                k = 0
                for j in range(len(dw)):
                    time_interval = []
                    time_interval1 = ''
                    time_interval2 = ''
                    if j < 137 and j > 28 and (j == 29 or (j - 29) % 3 == 0):
                        if math.isnan(dw[j][i]) and math.isnan(dw[j][i + 2]):
                            time_interval1 = None
                        else:
                            if dw[j][i+1] == ':':
                                time_interval1 = str(float(dw[j][i])+0.5) + '-' + str(float(dw[j][i + 2]))
                            elif dw[j][i+1] == ';':
                                time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            elif dw[j][i+1] == ':;':
                                time_interval1 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            else:
                                time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]))
                        if math.isnan(dw[j + 1][i]) and math.isnan(dw[j + 1][i + 2]):
                            time_interval2 = None
                        else:
                            if dw[j][i+1] == ':':
                                time_interval2 = str(float(dw[j][i])+ 0.5) + '-' + str(float(dw[j][i + 2]))
                            elif dw[j][i+1] == ';':
                                time_interval2 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            elif dw[j][i+1] == ':;':
                                time_interval2 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            else:
                                time_interval2 = str(float(dw[j + 1][i])) + '-' + str(float(dw[j + 1][i + 2]))
                        k += 1
                        time_interval.append(time_interval1)
                        time_interval.append(time_interval2)
                        oper_work[name[k - 1]] = time_interval
                    elif j < 184 and j > 137 and (j == 138 or (j - 138) % 3 == 0):
                        if math.isnan(dw[j][i]) and math.isnan(dw[j][i + 2]):
                            time_interval1 = None
                        else:
                            if dw[j][i+1] == ':':
                                time_interval1 = str(float(dw[j][i])+0.5) + '-' + str(float(dw[j][i + 2]))
                            elif dw[j][i+1] == ';':
                                time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            elif dw[j][i+1] == ':;':
                                time_interval1 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            else:
                                time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]))
                        if math.isnan(dw[j + 1][i]) and math.isnan(dw[j + 1][i + 2]):
                            time_interval2 = None
                        else:
                            if dw[j][i+1] == ':':
                                time_interval2 = str(float(dw[j][i])+0.5) + '-' + str(float(dw[j][i + 2]))
                            elif dw[j][i+1] == ';':
                                time_interval2 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            elif dw[j][i+1] == ':;':
                                time_interval2 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]) + 0.5)
                            else:
                                time_interval2 = str(float(dw[j + 1][i])) + '-' + str(float(dw[j + 1][i + 2]))
                        k += 1
                        time_interval.append(time_interval1)
                        time_interval.append(time_interval2)
                        oper_work[name[k - 1]] = time_interval
                date_data[day] = oper_work
                day += 1
        return name, date_data

    sheet_name = 'График работы(Ноябрь1)'
    month_number = {'Январь': 1, 'Февраль': 2, 'Март': 3, 'Апрель': 4, 'Май': 5, 'Июнь': 6, 'Июль': 7, 'Август': 8,
                    'Сентябрь': 9, 'Октябрь': 10, 'Ноябрь': 11, 'Декабрь': 12}
    month = sheet_name.split('(')[1][:-2]
    name, date_data = scan_excel(sheet_name, month_number[month])
    day_now = int(datetime.datetime.now().day) + day_plus
    day_work_report = 'График ' + str(day_now) + '.' + str(month_number[month]) + '.2024 \n\n\n'

    data_date_time = []
    data_date_night = []
    data_date_night_next = []
    first_time = 0
    for work_time in date_data[day_now]:
        if work_time not in night:
            if (date_data[day_now].get(work_time)[0] or date_data[day_now].get(work_time)[1]):
                if date_data[day_now].get(work_time)[0]:
                    k = []
                    k.append(float(date_data[day_now].get(work_time)[0].split('-')[0]))
                    k.append(float(date_data[day_now].get(work_time)[0].split('-')[1]))
                    k.append(date_data[day_now].get(work_time)[0])
                    k.append(work_time)
                    data_date_time.append(k)
                if date_data[day_now].get(work_time)[1]:
                    k = []
                    k.append(float(date_data[day_now].get(work_time)[1].split('-')[0]))
                    k.append(float(date_data[day_now].get(work_time)[1].split('-')[1]))
                    k.append(date_data[day_now].get(work_time)[1])
                    k.append(work_time)
                    data_date_time.append(k)
        # Для ночников
        else:
            if (date_data[day_now].get(work_time)[0] or date_data[day_now].get(work_time)[1]):
                if date_data[day_now].get(work_time)[0]:
                    k = []
                    k.append(float(date_data[day_now].get(work_time)[0].split('-')[0]))
                    k.append(float(date_data[day_now].get(work_time)[0].split('-')[1]))
                    k.append(date_data[day_now].get(work_time)[0])
                    k.append(work_time)
                    if k[0] >= 18:
                        data_date_night.append(k)
                if date_data[day_now].get(work_time)[1]:
                    k = []
                    k.append(float(date_data[day_now].get(work_time)[1].split('-')[0]))
                    k.append(float(date_data[day_now].get(work_time)[1].split('-')[1]))
                    k.append(date_data[day_now].get(work_time)[1])
                    k.append(work_time)
                    if k[0] >= 18:
                        data_date_night.append(k)
    #   Для ночников следующий день
    for work_time in date_data[day_now + 1]:
        if work_time in night:
            if (date_data[day_now + 1].get(work_time)[0] or date_data[day_now + 1].get(work_time)[1]):
                if date_data[day_now + 1].get(work_time)[0]:
                    k = []
                    k.append(float(date_data[day_now + 1].get(work_time)[0].split('-')[0]))
                    k.append(float(date_data[day_now + 1].get(work_time)[0].split('-')[1]))
                    k.append(date_data[day_now + 1].get(work_time)[0])
                    k.append(work_time)
                    if k[0] == 0:
                        data_date_night_next.append(k)
                if date_data[day_now + 1].get(work_time)[1]:
                    k = []
                    k.append(float(date_data[day_now + 1].get(work_time)[1].split('-')[0]))
                    k.append(float(date_data[day_now + 1].get(work_time)[1].split('-')[1]))
                    k.append(date_data[day_now + 1].get(work_time)[1])
                    k.append(work_time)
                    if k[0] == 0:
                        data_date_night_next.append(k)
    data_work_time_night = []
    for i in range(0, len(data_date_night)):
        k = []
        k.append(data_date_night[i][0])
        k.append(float(data_date_night_next[i][2].split('-')[1]))
        k.append(data_date_night[i][2].split('-')[0] + '-' + data_date_night_next[i][2].split('-')[1])
        k.append(data_date_night[i][3])
        data_work_time_night.append(k)
    # Фильтер для дневников
    worker_filter1 = []
    for i in range(0, 48):
        for worker in data_date_time:
            if worker[0] == i*0.5:
                worker_filter1.append(worker)
    for j in range(3):
        for i in range(0, len(worker_filter1)):
            if i < (len(worker_filter1)-1) and int(worker_filter1[i][0]) == int(worker_filter1[i+1][0]) and worker_filter1[i][1] > worker_filter1[i+1][1]:
                worker_filter1[i], worker_filter1[i+1] = worker_filter1[i+1], worker_filter1[i]
    for i in range(len(worker_filter1)):
        if worker_filter1[i][0]%1 !=0 and worker_filter1[i][1]%1 !=0:
            worker_filter1[i][2] = str(int(worker_filter1[i][0])) + ':30-' + str(int(worker_filter1[i][1])) + ':30'
        elif worker_filter1[i][0]%1 !=0 and worker_filter1[i][1]%1 ==0:
            worker_filter1[i][2] = str(int(worker_filter1[i][0])) + ':30-' + str(int(worker_filter1[i][1]))
        elif worker_filter1[i][0]%1 ==0 and worker_filter1[i][1]%1 !=0:
            worker_filter1[i][2] = str(int(worker_filter1[i][0])) + '-' + str(int(worker_filter1[i][1])) + ':30'
        else:
            worker_filter1[i][2] = str(int(worker_filter1[i][0])) + '-' + str(int(worker_filter1[i][1]))
    time_worker = {}
    for worker in worker_filter1:
        time_worker[worker[2]] = time_worker.get(worker[2], '') + worker[3] + '.'
    #     закончилось для дневников
    # Фильтер для ночников
    worker_filter2 = []
    for i in range(19, 24):
        for worker in data_work_time_night:
            if worker[0] == i:
                worker_filter2.append(worker)
    for j in range(3):
        for i in range(0, len(worker_filter2)):
            if i < (len(worker_filter2)-1) and int(worker_filter2[i][0]) == int(worker_filter2[i+1][0]) and worker_filter2[i][1] > worker_filter2[i+1][1]:
                worker_filter2[i], worker_filter2[i+1] = worker_filter2[i+1], worker_filter2[i]
    for i in range(len(worker_filter2)):
        worker_filter2[i][2] = str(int(worker_filter2[i][0])) + '-' + str(int(worker_filter2[i][1]))
    time_worker_night = {}
    for worker in worker_filter2:
        time_worker_night[worker[2]] = time_worker_night.get(worker[2], '') + worker[3] + '.'
    #     закончилось для ночников
    for worker in time_worker:
        first_name = str(time_worker[worker]).split('.')
        if first_name[1] == '':
            day_work_report += worker + '\n' + first_name[0] + ' ' + operators_tags.get(first_name[0], 'нет') + '\n\n'
        else:
            day_work_report += (worker + '\n' + first_name[0] + ' ' + operators_tags.get(first_name[0], 'нет') + '\n' +
                                first_name[1] + ' ' + operators_tags.get(first_name[1], 'нет') + '\n\n')
    for worker in time_worker_night:
        first_name = str(time_worker_night[worker]).split('.')
        day_work_report += worker + '\n' + first_name[0] + ' ' + operators_tags.get(first_name[0], 'нет') + '\n\n'
    day_work_report += '\n\n' + 'https://docs.google.com/spreadsheets/d/1MQ0hNCotH-ou1jp3ODHeBhdww1w6UhG8X588-n-db6c/edit?gid=1387237092#gid=1387237092'
    return day_work_report