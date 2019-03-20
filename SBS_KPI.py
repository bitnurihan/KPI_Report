import os
from openpyxl import load_workbook
import data as data
import pandas as pd
import xlrd

path_dir = r'C:\Users\hanbi01\Desktop\한빛누리\(분기)KPI\raw_data'

file_list = os.listdir(path_dir)

# Nuri = {"program_name" :  '그것이 알고싶다', 'copy_cells': [1, 2, 3, 4, 5]}
# print(Nuri['program_name'])


def delete_cols():
    global col
    for col in ['I', 'K', 'L', 'M', 'P', 'Q', 'R']:
        worksheet_read.column_dimensions[col].hidden = True


for file in file_list:
    raw_data_files = xlrd.open_workbook('%s\%s' % (path_dir, file))

    print(file)

worksheet_read = raw_data_files.sheet_by_name('sbs_kpi_')
delete_cols()