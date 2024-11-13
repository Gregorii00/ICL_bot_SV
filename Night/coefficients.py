import datetime
import csv
from additional import src_files
def coefficients(File_name):
    file = src_files + File_name

    with open(file, encoding='utf-8') as r_file:
        file_reader = csv.reader(r_file, delimiter=",")
        count = 0
        result = ''
        # Считывание данных из CSV файла
        for row in file_reader:
            if count >0:
                # Вывод строк
                if len(row[3])>5:
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
                hour = int(h)
                min = ((int(h) * 3600 + int(m) * 60 + int(s))-int(h)*3600)/60
                index = '-0'
                res = hour
                if min<5:
                    res = float(res)
                if min>=5 and min<15:
                    res +=0.2
                    index = '-0,83'
                if min>=15 and min<25:
                    res +=0.3
                    index = '-0,67'
                if min>=25 and min<35:
                    res +=0.5
                    index = '-0,5'
                if min>=35 and min<45:
                    res +=0.7
                    index = '-0,33'
                if min>=45 and min<55:
                    res +=0.8
                    index = '-0,17'
                if min>=55:
                    res = float(res+1)
                res_str = f'{row[0]}:   {res}'
                s=(73 - (len(res_str) + len(str(index))))
                for k in range((73 - (len(res_str) + len(str(index))))):
                    res_str+=' '
                res_str+=index +'\n'
                result+=res_str
            count += 1
        print(f'Всего в файле {count} строк.')
    return result