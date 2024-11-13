import datetime
from additional import report_week_name1, report_week_name2
date_now = datetime.datetime.now()
date_start = date_now - datetime.timedelta(days=7)
date_last = date_now - datetime.timedelta(days=1)
str_date_start = str(date_start.date().day) + '.' + str(date_start.date().month)
str_date_last = str(date_last.date().day) + '.' + str(date_last.date().month) + '.' + str(date_last.date().year)
date_name = str_date_start + '-' + str_date_last

name_1 = report_week_name1 + date_name
name_2 = report_week_name2 + date_name
print(name_1)
print(name_2)