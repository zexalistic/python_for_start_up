import os
import pandas as pd
import xlrd
import datetime

raw_file = 'Export Tracking 截止目前最新.xls'
content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
df_raw = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online

sheet_name = 'Vendor Export 2021'
data = df_raw.get(sheet_name).values
subject = data[2]

start_time = datetime.datetime(2021, 6, 30, 23, 59, 59)
end_time = datetime.datetime(2021, 7, 31, 23, 59, 59)

start_idx = 0
for i in range(3, len(data)):
    if isinstance(data[i][2], datetime.datetime) and data[i][2] > start_time:
        start_idx = i
        break

end_idx = len(data)
for i in range(start_idx, len(data)):
    if isinstance(data[i][2], datetime.datetime) and data[i][2] > end_time:
        end_idx = i
        break

todo = data[start_idx:end_idx]

sale_time_list = list()
client_name_list = list()
amount_usd_list = list()
exchange_rate = 6.47
sale_time = start_time
for record in todo:
    if pd.notna(record[2]):
        sale_time = record[2]
    client_name = record[4]
    amount = record[5]
    if isinstance(sale_time, datetime.datetime):
        sale_time_list.append(sale_time.date())
    else:
        sale_time_list.append(sale_time)
    client_name_list.append(client_name)
    amount_usd_list.append(amount)

export_dict = dict()
len_of_record = len(client_name_list)
export_dict['凭证类别'] = ['记'] * len_of_record
export_dict['凭证号'] = ['003'] * len_of_record
export_dict['凭证日期'] = sale_time_list
export_dict['附单据数'] = [None] * len_of_record
export_dict['摘要'] = ['采购成本'] * len_of_record
export_dict['科目编码'] = ['220203'] * len_of_record
export_dict['借方金额'] = [None] * len_of_record
export_dict['贷方金额'] = [i * exchange_rate for i in amount_usd_list]
export_dict['供应商'] = client_name_list
export_dict['外币金额'] = amount_usd_list
export_dict['币种'] = ['USD'] * len_of_record
export_dict['汇率'] = [exchange_rate] * len_of_record
export_dict['制单人'] = ['Daisy'] * len_of_record
export_dict['审核人'] = ['Sara'] * len_of_record

"""
Write to new excel
"""
template = 'template.xlsx'
df_export = pd.read_excel(template, engine='openpyxl', sheet_name=None)             # copied online

# sheet = df_export.get('导入-采购成本')
# columns = sheet.columns.values
new_sheet = pd.DataFrame(export_dict)
new_sheet.to_excel('output.xlsx', 'name of sheet')



