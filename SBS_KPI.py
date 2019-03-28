import pyexcel
from openpyxl import load_workbook
import os
from collections import Counter

# *.xls -> *.xlsx
path_dir = r'C:\Users\hanbi01\Desktop\한빛누리\(분기)KPI\raw_data'
file_list = os.listdir(path_dir)
new_path_dir = r'C:\Users\hanbi01\Desktop\한빛누리\(분기)KPI\xlsx_raw_data'
new_file_list = os.listdir(new_path_dir)

for file in file_list:
    pyexcel.save_book_as(file_name=("%s\%s" % (path_dir, file)), dest_file_name="%s\%s.xlsx" % (new_path_dir, file.split('.')[0]))

# 쓸모없는 열 지우기
for file in new_file_list:
    new_file = load_workbook("%s\%s" % (new_path_dir, file))

    if os.path.splitext(file)[0] == '1_10': #1_11/ 1_12 파일은 해당 열을 지울 필요가 없음. pass 처리
        pass
    elif os.path.splitext(file)[0] == '1_11':
        pass
    else:
        sheet = new_file['전체 수도권']
        sheet.delete_cols(18)
        sheet.delete_cols(17)
        sheet.delete_cols(16)
        sheet.delete_cols(13)
        sheet.delete_cols(12)
        sheet.delete_cols(11)
        sheet.delete_cols(9)
        new_file.save("%s" % file)


# 파일 내용 복사하기
