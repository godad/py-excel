# -*- coding: utf-8 -*-
# @Time    : 2018/3/13 13:35
# @Author  : flyfish
# @Email   : im@flyfish.im

# import openpyxl
# from openpyxl.workbook import Workbook

import os
import pyexcel as pe
import string

# def main(base_dir):
#     # Simple give the file name to **Reader**
#     # "example.xls","example.xlsx","example.ods", "example.xlsm"
#     spreadsheet = pe.get_sheet(file_name=os.path.join(base_dir, "example.xlsx"))
#
#     # row_range() gives [0 .. number of rows]
#     for r in spreadsheet.row_range():
#         # column_range() gives [0 .. number of ranges]
#         for c in spreadsheet.column_range():
#             # cell_value(row_index, column_index)
#             # return the value at the specified
#             # position
#             # please note that both row_index
#             # and column_index starts from 0
#             print(spreadsheet.cell_value(r, c))

def handle_excel():
    records = pe.get_records(file_name="example.xlsx")
    for record in records:
        print("%s 剩余应收款为 %d" % (record['销售名称'], record['剩余应收款']))
        if int(record['剩余应收款']) > 0 :
            print("正在将本条单独拆分出来...")
        elif str.find( record['销售名称'], '\/' ):
            # if '\/' in record['销售名称']:
                print("有两个销售")
        else:
            print("剩余应收款为0，不做操作...")
        print("表格拆分完毕，即将进入下一步，发送邮件")

if __name__ == '__main__':
        handle_excel()
# def handle_excel():
#     wb = openpyxl.load_workbook("2018年应收款表模板 - test.xlsx")
#     #print(wb.sheetnames)
#     ws = wb['Sheet1']
#     #print(ws.max_row)
#    for row in ws.iter_cols(min_row=2, max_row=8,min_col=1,max_col=24):
#     #for column in ws.iter_rows(min_row=3, max_row=8, min_col=1, max_col=24):
#     print(row)
#        for cell in excel:
#        if column[18].value > 0:
#            print(column[0].value,column[1].value,column[2].value,column[3].value,column[6].value,column[9].value,column[12].value,column[13].value,column[14].value,column[15].value,column[16].value,column[18].value,column[19].value,column[20].value)
#          outwb = Workbook()
#         # 第一个sheet是ws
#         # outws = wb.worksheets[0]
#
#          outws.append()
#          outwb.save("test.xlsx")
# def create_write_to_workbook():
#     wb = Workbook()
#
#     ws1 = wb.active
#     ws1.title = 'Sheet1'
#
#     ws1['A1'] = '销售名称'
#     ws1['B1'] = '合同编号'
#     ws1['C1'] = '客户名称'
#     ws1['D1'] = '出库状态'
#     ws1['E1'] = '开票状态'
#     ws1['F1'] = '应收阶段'
#     ws1['G1'] = '应收阶段'
#     ws1['H1'] = '触发应收条款'
#     ws1['I1'] = '到期应收金额'
#     ws1['J1'] = '逾期天数(D)'
#     ws1['K1'] = '剩余应收款'
#     ws1['L1'] = '计划回款时间'
#
#     wb.save("test1.xlsx")






# if __name__ == "__main__":
#     create_write_to_workbook()