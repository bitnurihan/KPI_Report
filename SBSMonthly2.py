from datetime import datetime, date, timedelta
from openpyxl import load_workbook
from copy import copy
import xlrd



def read_data_from_excel(write_file_name):
    global read_excel_file, worksheet_read, month, row, col, array, inner_array, i, j, value
    read_excel_file = xlrd.open_workbook(write_file_name)
    worksheet_read = read_excel_file.sheet_by_name('전체 수도권 (P) - 가구 (1408)')
    row = 3
    col = 4
    array = []
    inner_array = []
    for i in range(1):
        for j in range(10):
            value = worksheet_read.cell_value(rowx=row, colx=col)
            inner_array.append(value)
            col += 1

        array.append(inner_array)
        inner_array = []
        col = 3
        row += 1
        print()


def paste_to_excel(row, col):
    global i, j, test, test
    for i in range(1):
        for j in range(10):
            test = worksheet_write.cell(row=row, column=col)
            test.value = array[i][j]
            col += 1
        print()


###1. 기존전시간대시청률(06-11,17-24)
write_excel_file = load_workbook(filename=r'C:\Users\hanbi01\Desktop\한빛누리\(매월)SBS월간업데이트\MonthlyReport2.xlsx')

read_data_from_excel(r'C:\Users\hanbi01\Desktop\한빛누리\(매월)SBS월간업데이트\1.xls')
worksheet_write = write_excel_file[r'기존전시간대시청률(06-11,17-24)']

paste_to_excel(380, 3)  ## 380이 매달 +2씩 증가할 수 있게 수정해야 함.


### 셀 스타일 복사해서 붙여넣기

row = 378 ## 378이 매달 +2씩 증가할 수 있게 수정해야 함.
col = 1

for i in range(2):
    for j in range(16):
        worksheet_write.cell(row+2, col)._style = copy(worksheet_write.cell(row, col)._style)
        col += 1

    row += 1
    col = 1

## 함수 복붙하기


row = 378 ## 378이 매달 +2씩 증가할 수 있게 수정해야 함.
col = 1

for i in range(1):
    for j in range(16):
        worksheet_write.cell(row+3, col).value = copy(worksheet_write.cell(row+1, col).value) # 이렇게 했더니 379번줄 그대로 가져와서 문제생김.. -_-
        col += 1


### 셀숨기기
for col in ['N', 'O', 'P']:
    worksheet_write.column_dimensions[col].hidden = True

## 날짜넣기

row = 380 ## 380이 매달 +2씩 증가할 수 있게 수정해야 함.
month = datetime.now().month

year = datetime.now().year
date_cell = worksheet_write.cell(row=row, column=1)
date_cell.value = "%s. %s" % (year, (month-1))


write_excel_file.save('testfile2.xlsx')
