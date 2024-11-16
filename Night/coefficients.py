import datetime
import csv
from additional import src_files
from Night.search_exel_month import scan_excel
def coefficients(File_name):
    file = src_files + File_name
    excel_name_coef = {}
    with open(file, encoding='utf-8') as r_file:
        sv_name = ['Егоров Олег Александрович', 'Жевнер Григорий Павлович', 'Павликов Илья Сергеевич',
                   'Янцевич Маргарита Александровна', 'Киселева Дарина Максимовна']
        file_reader = csv.reader(r_file, delimiter=",")
        day_i = ''
        count = 0
        work_time_name = {}
        str_result = []
        for row in file_reader:
            # if count >0 and row[0] not in sv_name:
            if count >0:
                if len(row[3]) > 5:
                    (h1, m1, s1) = row[3].split(':')
                else:
                    (m1, s1) = row[3].split(':')
                    h1='0'
                if len(row[9])>5:
                    (h2, m2, s2) = row[9].split(':')
                else:
                    (m2, s2) = row[9].split(':')
                    h2='0'
                mysum = datetime.timedelta()
                d1 = datetime.timedelta(hours=int(h1), minutes=int(m1), seconds=int(s1))
                mysum+=d1
                d2 = datetime.timedelta(hours=int(h2), minutes=int(m2), seconds=int(s2))
                mysum += d2
                (h, m, s) = str(mysum).split(':')
                hour = float(h)
                min = ((int(h) * 3600 + int(m) * 60 + int(s))-int(h)*3600)/60
                work_time_name[row[0]] = [hour, round(min, 2)]
                day_i = row[10].split('.')[0]
            count+=1
        name, date_data, month = scan_excel()
        print(date_data)
        result = f' Коэффициенты за {day_i}.{month} \n\n'
        day_now = int(day_i)
        data_date_time = []
        data_work_time_day = {}
        for work_time in date_data[day_now]:
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
        for work in data_date_time:
            data_work_time_day[work[3]] = data_work_time_day.get(work[3], 0) + work[1] - work[0]
        for work in work_time_name:
            res_str = ''
            if work in name:
                if int(data_work_time_day[work] - work_time_name[work][0]) >0:
                    res = work_time_name[work][0]
                    if work_time_name[work][1] < 5:
                        res = float(res)
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0]-1)) + ',00'
                    if work_time_name[work][1] >= 5 and work_time_name[work][1] < 15:
                        res += 0.2
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0]-1))+ ',83'
                    if work_time_name[work][1] >= 15 and work_time_name[work][1] < 25:
                        res += 0.3
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0]-1)) + ',67'
                    if work_time_name[work][1] >= 25 and work_time_name[work][1] < 35:
                        res += 0.5
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0]-1)) + ',5'
                    if work_time_name[work][1] >= 35 and work_time_name[work][1] < 45:
                        res += 0.7
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0]-1)) + ',33'
                    if work_time_name[work][1] >= 45 and work_time_name[work][1] < 55:
                        res += 0.8
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0]-1)) + ',17'
                    if work_time_name[work][1] >= 55:
                        res = float(res + 1)
                        index = '-' + str(int(data_work_time_day[work] - work_time_name[work][0])-1) + ',00'
                    excel_name_coef[work] = index
                    space1 = int(30 - len(str(work)))
                    space2 = 10
                    res_str = str(work)+ ':' + ('⠀'* space1) + str(res) + '⠀'* space2 + str(index) + '\n'
                    str_result.append(res_str)
                elif (data_work_time_day[work] - work_time_name[work][0]) <= 0:
                    res = work_time_name[work][0]
                    if work_time_name[work][1] < 5:
                        res = float(res)
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0])) + ',00'
                    if work_time_name[work][1] >= 5 and work_time_name[work][1] < 15:
                        res += 0.2
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0]))+ ',17'
                    if work_time_name[work][1] >= 15 and work_time_name[work][1] < 25:
                        res += 0.3
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0])) + ',33'
                    if work_time_name[work][1] >= 25 and work_time_name[work][1] < 35:
                        res += 0.5
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0])) + ',5'
                    if work_time_name[work][1] >= 35 and work_time_name[work][1] < 45:
                        res += 0.7
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0])) + ',67'
                    if work_time_name[work][1] >= 45 and work_time_name[work][1] < 55:
                        res += 0.8
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0])) + ',88'
                    if work_time_name[work][1] >= 55:
                        res = float(res + 1)
                        index = '+' + str(int(data_work_time_day[work] - work_time_name[work][0])) + ',00'
                    excel_name_coef[work] = index
                    space1 = int(30 - len(str(work)))
                    space2 = 10
                    res_str = str(work)+ ':' + ('⠀' * space1) + str(res) + '⠀'* space2 + str(index) + '\n'
                    str_result.append(res_str)
            else:
                res = work_time_name[work][0]
                if work_time_name[work][1] < 5:
                    res = float(res)
                    index = '-0,00'
                if work_time_name[work][1] >= 5 and work_time_name[work][1] < 15:
                    res += 0.2
                    index = '-0,83'
                if work_time_name[work][1] >= 15 and work_time_name[work][1] < 25:
                    res += 0.3
                    index = '-0,67'
                if work_time_name[work][1] >= 25 and work_time_name[work][1] < 35:
                    res += 0.5
                    index = '-0,5'
                if work_time_name[work][1] >= 35 and work_time_name[work][1] < 45:
                    res += 0.7
                    index = '-0,33'
                if work_time_name[work][1] >= 45 and work_time_name[work][1] < 55:
                    res += 0.8
                    index = '-0,17'
                if work_time_name[work][1] >= 55:
                    res = float(res + 1)
                    index = '-0,00'
                excel_name_coef[work] = index
                space1 = int(30 - len(str(work)))
                space2  = 15 - len(index)
                res_str = str(work) + ':' + ('⠀' * space1) + str(res) + '⠀'* space2 + str(index) + '\n'
                str_result.append(res_str)
        a_count = 0
        for a in str_result:
            a_count+=len(a)
            result+=a
    return result, excel_name_coef, day_now
