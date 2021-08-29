import os
import pandas as pd
import xlrd
import datetime
import logging

# template = 'template.xlsx'
# df_export = pd.read_excel(template, engine='openpyxl', sheet_name=None)             # copied online

raw_file = 'Export Tracking 截止目前最新.xls'
content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
df_raw = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online

sheet_name = 'Vendor Export 2021'
data = df_raw.get(sheet_name).values
subject = data[2]

"""
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
"""


# sale_time_list = list()
# client_name_list = list()
# amount_usd_list = list()
# exchange_rate = 6.47
# sale_time = start_time
# for record in todo:
#     if pd.notna(record[2]):
#         sale_time = record[2]
#     client_name = record[4]
#     amount = record[5]
#     if isinstance(sale_time, datetime.datetime):
#         sale_time_list.append(sale_time.date())
#     else:
#         sale_time_list.append(sale_time)
#     client_name_list.append(client_name)
#     amount_usd_list.append(amount)

raw_file = '导入模板-数据生成.xls'
content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
abo = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online

sheet_name = '供应商编号'
gongyinshang = abo.get(sheet_name).values

gys = dict()
for rec in gongyinshang:
    if pd.notna(rec[1]):
        gys[rec[1].upper()] = rec[2]

month = '17月'
todo = list()
for i in range(3, len(data)):
    if data[i][0] == month and pd.notna(data[i][17]):
        todo.append(data[i])

sale_time_list = list()
client_name_list = list()
amount_usd_list = list()
exchange_rate = 6.4728
sale_time = month
invoice_num_list = list()
for record in todo:
    client_name = record[15]
    amount = record[17]
    invoice_num_list.append(record[13])
    sale_time_list.append(sale_time)
    if gys.__contains__(client_name.upper()):
        client_name = gys[client_name.upper()]
    else:
        logging.error(f'Can not find client name: {client_name}')
    client_name_list.append(client_name)
    amount_usd_list.append(round(amount, 2))

export_dict = dict()
len_of_record = len(client_name_list) + 1
export_dict['凭证类别'] = ['记'] * len_of_record
export_dict['凭证号'] = ['003'] * len_of_record
export_dict['凭证日期'] = ['2021-7-31'] * len_of_record
export_dict['附单据数'] = [None] * len_of_record
export_dict['摘要'] = ['采购成本'] + ['采购成本 ' + str(j) + ' ' + str(i) for i, j in zip(invoice_num_list, client_name_list)]
export_dict['科目编码'] = ['140501'] + ['220203'] * (len_of_record-1)
export_dict['借方金额'] = [round(sum(amount_usd_list) * exchange_rate, 2)] + [None] * (len_of_record - 1)
export_dict['贷方金额'] = [None] + [round(i * exchange_rate, 2) for i in amount_usd_list]
export_dict['客户'] = [None] * len_of_record
export_dict['供应商'] = [None] + client_name_list
export_dict['外币金额'] = [None] + amount_usd_list
export_dict['币种'] = ['USD'] * len_of_record
export_dict['汇率'] = [exchange_rate] * len_of_record
export_dict['制单人'] = ['Daisy'] * len_of_record
export_dict['审核人'] = ['Sara'] * len_of_record

"""
Write to new excel
"""

# sheet = df_export.get('导入-采购成本')
# columns = sheet.columns.values
new_sheet = pd.DataFrame(export_dict)
new_sheet.to_excel('output.xlsx', 'name of sheet')



