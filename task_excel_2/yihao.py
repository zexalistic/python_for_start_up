import os
import pandas as pd
import xlrd

# raw_file = 'Export Tracking 截止目前最新.xls'
# content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
# df_raw = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online

# sheet_name = 'Vendor Export 2021'
# data = df_raw.get(sheet_name).values
# subject = data[2]
#
# # for record in range(3, len(data)):
# #     month = data[record]
# export = data[3][2]


"""
Write to new excel
"""
template = 'template.xlsx'
df_export = pd.read_excel(template, engine='openpyxl', sheet_name=None)             # copied online

sheet = df_export.get('导入-采购成本')
columns = sheet.columns.values
df_new_data = pd.DataFrame(columns, [1] * len(columns))
sheet.append(df_new_data)
sheet.to_excel('output.xlsx', 'name of sheet')



