import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from additional import scope, credentials_file, spreadsheet_id_wiretapping
def excel_scan_wiretapping(start, end, report=True):
    if report:
        credentials = ServiceAccountCredentials.from_json_keyfile_name('../'+credentials_file, scopes=scope)
    else:
        credentials = ServiceAccountCredentials.from_json_keyfile_name('./'+credentials_file, scopes=scope)
    print(start)
    print(end)
    client = gspread.authorize(credentials)
    sheet = client.open_by_key(spreadsheet_id_wiretapping)
    file1 = sheet.get_worksheet(4)
    data = file1.get_all_values()
    data_result = []
    for i in range(6719, file1.row_count):
        k = []
        if data[i][0] != '' and (data[i][0].split(' ')[1]!=''):
            date = data[i][0].split(' ')[0].split('.')
            if date[1] == str(start).split('-')[1] and (int(date[0]) >=int(str(start).split('-')[0]) and int(date[0]) <=int(str(end).split('-')[0])) and data[i][6] == 'Необоснованная':
                k.append(data[i][3].strip())
                k.append(int(data[i][2]))
                data_result.append(k)
    return data_result