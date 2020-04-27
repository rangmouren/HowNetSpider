# !/usr/bin/env python
# -*-coding: utf-8-*-
import xlsxwriter


# 生成excel文件
def generate_excel(rec_data,name):
    workbook = xlsxwriter.Workbook('./file/{}.xlsx'.format(name))
    worksheet = workbook.add_worksheet()

     # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
     # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})
    # money_format = workbook.add_format({'num_format': '$#,##0'})
    # date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

     # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15)

     # 用符号标记位置，例如：A列1行
    worksheet.write('A1', 'title', bold_format)
    worksheet.write('B1', 'url', bold_format)
    worksheet.write('C1', 'authors', bold_format)
    worksheet.write('D1', 'source', bold_format)
    worksheet.write('E1', 'degree', bold_format)
    worksheet.write('F1', 'time', bold_format)

    row = 1
    col = 0
    for item in (rec_data):
        # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, item['title'])
        worksheet.write_string(row, col + 1, item['url'])
        worksheet.write_string(row, col + 2, item['authors'])
        worksheet.write_string(row, col + 3, item['source'])
        worksheet.write_string(row, col + 4, item['degree'])
        worksheet.write_string(row, col + 5, item['time'])
        row += 1
    workbook.close()

    # coding=utf-8



# 读取execl
import xlrd


def read_xlrd(excelFile):
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    dataFile = []

    for rowNum in range(table.nrows):
        # if 去掉表头
        if rowNum > 0:
            dataFile.append(table.row_values(rowNum))

    return dataFile


if __name__ == '__main__':
    excelFile = 'file/demo.xlsx'
    print(read_xlrd(excelFile=excelFile))