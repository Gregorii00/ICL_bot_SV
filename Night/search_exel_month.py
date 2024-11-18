import datetime
import math
from calendar import monthrange
import pandas as pd
from additional import spreadsheet_id_schedule
def scan_excel(month_now_print = 0):
    spreadsheet_id = spreadsheet_id_schedule

    url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet_id + '/export?format=xlsx&rtpof=true&sd=true'
    sheet_name1 = 'График работы('
    sheet_name2 = '1)'
    month_number = {1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
                    9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'}
    if month_now_print == 0:
        date_now = datetime.datetime.now()
        month_now = date_now.month
    else:
        month_now = month_now_print
    sheet_name = sheet_name1 + month_number[month_now] + sheet_name2
    now_year = datetime.datetime.now().year
    days = monthrange(now_year, month_now)[1]

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
                if j < 138 and j > 28 and (j == 30 or (j - 30) % 3 == 0):
                    name.append(dw[j][i])
                elif j < 185 and j > 138 and (j == 139 or (j - 139) % 3 == 0):
                    name.append(dw[j][i])
        elif (i - 3) % 3 == 0:
            k = 0
            for j in range(len(dw)):
                time_interval = []
                time_interval1 = ''
                time_interval2 = ''
                if j < 138 and j > 29 and (j == 30 or (j - 30) % 3 == 0):
                    if math.isnan(dw[j][i]) and math.isnan(dw[j][i + 2]):
                        time_interval1 = None
                    else:
                        if dw[j][i + 1] == ':':
                            time_interval1 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]))
                        elif dw[j][i + 1] == ';':
                            time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]) + 0.5)
                        elif dw[j][i + 1] == ':;':
                            time_interval1 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]) + 0.5)
                        else:
                            time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]))
                    if math.isnan(dw[j + 1][i]) and math.isnan(dw[j + 1][i + 2]):
                        time_interval2 = None
                    else:
                        if dw[j + 1][i + 1] == ':':
                            time_interval2 = str(float(dw[j + 1][i]) + 0.5) + '-' + str(float(dw[j + 1][i + 2]))
                        elif dw[j + 1][i + 1] == ';':
                            time_interval2 = str(float(dw[j + 1][i])) + '-' + str(float(dw[j + 1][i + 2]) + 0.5)
                        elif dw[j + 1][i + 1] == ':;':
                            time_interval2 = str(float(dw[j + 1][i]) + 0.5) + '-' + str(float(dw[j + 1][i + 2]) + 0.5)
                        else:
                            time_interval2 = str(float(dw[j + 1][i])) + '-' + str(float(dw[j + 1][i + 2]))
                    k += 1
                    time_interval.append(time_interval1)
                    time_interval.append(time_interval2)
                    oper_work[name[k - 1]] = time_interval
                elif j < 185 and j > 138 and (j == 139 or (j - 139) % 3 == 0):
                    if math.isnan(dw[j][i]) and math.isnan(dw[j][i + 2]):
                        time_interval1 = None
                    else:
                        if dw[j + 1][i + 1] == ':':
                            time_interval1 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]))
                        elif dw[j + 1][i + 1] == ';':
                            time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]) + 0.5)
                        elif dw[j + 1][i + 1] == ':;':
                            time_interval1 = str(float(dw[j][i]) + 0.5) + '-' + str(float(dw[j][i + 2]) + 0.5)
                        else:
                            time_interval1 = str(float(dw[j][i])) + '-' + str(float(dw[j][i + 2]))
                    if math.isnan(dw[j + 1][i]) and math.isnan(dw[j + 1][i + 2]):
                        time_interval2 = None
                    else:
                        if dw[j][i + 1] == ':':
                            time_interval2 = str(float(dw[j + 1][i]) + 0.5) + '-' + str(float(dw[j + 1][i + 2]))
                        elif dw[j][i + 1] == ';':
                            time_interval2 = str(float(dw[j + 1][i])) + '-' + str(float(dw[j + 1][i + 2]) + 0.5)
                        elif dw[j][i + 1] == ':;':
                            time_interval2 = str(float(dw[j + 1][i]) + 0.5) + '-' + str(float(dw[j + 1][i + 2]) + 0.5)
                        else:
                            time_interval2 = str(float(dw[j + 1][i])) + '-' + str(float(dw[j + 1][i + 2]))
                    k += 1
                    time_interval.append(time_interval1)
                    time_interval.append(time_interval2)
                    oper_work[name[k - 1]] = time_interval
            date_data[day] = oper_work
            day += 1
    return name, date_data, month_now