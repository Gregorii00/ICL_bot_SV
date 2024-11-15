import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from additional import scope, credentials_file, spreadsheet_id_wiretapping
def excel_scan_wiretapping():
    credentials = ServiceAccountCredentials.from_json_keyfile_name('../'+credentials_file, scopes=scope)
    client = gspread.authorize(credentials)
    sheet = client.open_by_key(spreadsheet_id_wiretapping)
    file1 = sheet.get_worksheet(4)
    data = file1.get_all_values()
    date_now = datetime.datetime.now()
    start = date_now.date() - datetime.timedelta(7)
    end = date_now.date() - datetime.timedelta(1)
    print(start)
    print(end)
    for i in range(6719, file1.row_count):
        if data[i][0] != '' and (data[i][0].split(' ')[1]!=''):
            date = data[i][0].split(' ')[0].split('.')
            if int(date[1]) == start.month and (int(date[0]) >=start.day and int(date[0]) <=end.day) and data[i][6] == 'Необоснованная':
                print(data[i])
    return 1

excel_scan_wiretapping()
