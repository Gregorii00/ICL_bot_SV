from Night.student_scan import excel_scan_student
from Night.coefficients import coefficients
from Night.excel_scan_wiretapping import excel_scan_wiretapping

files_name = ['Производительность_операторов_24.csv']
result, coef_name, day_now = coefficients(files_name[0], file_src_bool=False)
print('day_now: ', day_now)
student_result = excel_scan_student(day_now, credentials_bool=False)
print('student_result test2: ', student_result)
name_cef_in_formul = excel_scan_wiretapping(coef_name, day_now, student_result, credentials_bool=False)
print('name_cef_in_formul: ', name_cef_in_formul)



