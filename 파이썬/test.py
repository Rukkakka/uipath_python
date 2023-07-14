import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import shutil
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, GradientFill
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment

from excel import *



# df = pd.read_excel('통합 문서1.xlsx', 'Sheet1',header=0)

# print(df)

# for r in dataframe_to_rows(df, index=False, header=False):
#   ws.append(r)

# culumns is passed by list and element of columns means column index in worksheet.
# if culumns = [1, 3, 4] then, 1st, 3th, 4th columns are applied autofit culumn.
# margin is additional space of autofit column. 

## 각 칼럼에 대해서 모든 셀값의 문자열 개수에서 1.1만큼 곱한 것들 중 최대값을 계산한다.
# for column_cells in ws.columns:
#     length = max(len(str(cell.value))*1.1 for cell in column_cells)
#     ws.column_dimensions[column_cells[0].column_letter].width = length
#     ## 셀 가운데 정렬
#     for cell in ws[column_cells[0].column_letter]:
#         cell.alignment = Alignment(horizontal='center')
    
# wb.save('통합 문서1.xlsx')

append_range('통합 문서1.xlsx','Sheet1','통합 문서1.xlsx','Sheet2')



