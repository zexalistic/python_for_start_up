import os
import pandas as pd
import xlrd
import datetime
import logging


"""
Generate vendor English name to Chinese name dictionary
"""
raw_file = '导入模板-数据生成.xls'
content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
abo = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online
sheet_name = '供应商编号'
gongyinshang = abo.get(sheet_name).values

gys = dict()
for rec in gongyinshang:
    if pd.notna(rec[1]):
        gys[rec[1].upper()] = rec[2]

"""
Read Main Excel sheet
"""
# template = 'template.xlsx'
# df_export = pd.read_excel(template, engine='openpyxl', sheet_name=None)             # copied online

raw_file = 'Export Tracking 截止目前最新.xls'
content = xlrd.open_workbook(filename=raw_file, encoding_override='gbk')    # copied online
df_raw = pd.read_excel(content, engine='xlrd', sheet_name=None)             # copied online
sheet_name = 'Vendor Export 2021'
data = df_raw.get(sheet_name).values

month = '17月'
todo = list()
for i in range(3, len(data)):
    if data[i][0] == month and pd.notna(data[i][17]):
        todo.append(data[i])

exchange_rate = 6.4728

"""
Data processing
"""

class Deal:
    """
    A class representing the information of each deal we made
    """
    def __init__(self):
        self.client_name_list = list()
        self.amount_usd_list = list()
        self.order_num_list = list()
        self.ke_mu_bian_ma_jie = '140501'
        self.ke_mu_bian_ma_dai = '220203'


deal_list = list()
deal = Deal()
for record in todo:
    if pd.notna(record[2]):
        deal_list.append(deal)
        deal = Deal()
    client_name = record[15]
    if gys.__contains__(client_name.upper()):
        client_name = gys[client_name.upper()]
    else:
        logging.error(f'Can not find client name: {client_name}')
    deal.client_name_list.append(client_name)
    deal.amount_usd_list.append(record[17])
    deal.order_num_list.append(record[13])
deal_list.append(deal)
deal_list.pop(0)


kemubianma = list()
zhaiyao = list()
jiefangjine = list()
daifangjine = list()
len_of_record = len(deal_list)
gongyinshang = list()
waibijine = list()
for deal in deal_list:
    len_of_record += len(deal.client_name_list)
    kemubianma.append(deal.ke_mu_bian_ma_jie)
    zhaiyao.append('采购成本')
    jiefangjine.append(round(sum(deal.amount_usd_list) * exchange_rate, 2))
    daifangjine.append(None)
    gongyinshang.append(None)
    waibijine.append(None)
    for client_name, order_num, amount_usd in zip(deal.client_name_list, deal.order_num_list, deal.amount_usd_list):
        kemubianma.append(deal.ke_mu_bian_ma_dai)
        zhaiyao.append('采购成本 ' + str(client_name) + ' ' + str(order_num))
        jiefangjine.append(None)
        daifangjine.append(round(amount_usd * exchange_rate, 2))
        gongyinshang.append(client_name)
        waibijine.append(round(amount_usd, 2))



"""
Prepare to write to excel
"""

export_dict = dict()
export_dict['凭证类别'] = ['记'] * len_of_record
export_dict['凭证号'] = ['003'] * len_of_record
export_dict['凭证日期'] = ['2021-7-31'] * len_of_record
export_dict['附单据数'] = [None] * len_of_record
export_dict['摘要'] = zhaiyao
export_dict['科目编码'] = kemubianma
export_dict['借方金额'] = jiefangjine
export_dict['贷方金额'] = daifangjine
export_dict['客户'] = [None] * len_of_record
export_dict['供应商'] = gongyinshang
export_dict['外币金额'] = waibijine
export_dict['币种'] = ['USD'] * len_of_record
export_dict['汇率'] = [exchange_rate] * len_of_record
export_dict['制单人'] = ['Daisy'] * len_of_record
export_dict['审核人'] = ['Sara'] * len_of_record


# sheet = df_export.get('导入-采购成本')
# columns = sheet.columns.values
new_sheet = pd.DataFrame(export_dict)
new_sheet.to_excel('output_jenny.xlsx', 'name of sheet')



