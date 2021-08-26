"""
Here are some packages
"""
import os
import pandas as pd
import xlrd

"""
Iterate all files ends with xls
"""
for file in os.listdir():           # os.listdir() returns a list of files in current folder/ what is a list/ method/ loop
    if file.endswith('xls'):        # indent
        obj = file

raw_file = '利润表.xls'
content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
df_raw = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online
data = None                                                                 # Announce a variable
for sheet_name in df_raw:
    # print(sheet_name)                                                       # What is print
    sheet = df_raw.get(sheet_name)
    data = sheet.values                                                     # This is a list
    # print(data)

print(data[0])
print(data[0][0])

"""
An example calculating the sum of all tables
"""
sum = 0
for item in data:
    for element in item:
        if pd.isna(element):
            continue
        if isinstance(element, int) or isinstance(element, float):
            sum = sum + element

print(sum)