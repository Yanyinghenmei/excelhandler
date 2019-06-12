import xlrd
from datetime import date, datetime
import logging

file = 'test.xlsx'

def read_excel():
    wb = xlrd.open_workbook(filename=file)

    # 获取所有表格名字
    # print(wb.sheet_names())

    sheet1: xlrd.sheet.Sheet
    sheet2: xlrd.sheet.Sheet

    # 通过索引获取表格
    sheet1 = wb.sheet_by_index(0)
    # 通过名称获取表格
    sheet2 = wb.sheet_by_name('iOS端排期')

    # print(sheet1, sheet2)

    rows = sheet1.nrows
    cols = sheet1.ncols

    for row in range(0, rows):
        for col in range(0, cols):
            cell = sheet1.cell(row,col)
            print(cell.ctype, cell.value)

    # for p in rows:
    #     if isinstance(p,float) or isinstance(p,int):
    #         print(p)
    #     if isinstance(p,str):
    #         print(p.strip())
    #     if p.ctype == 3:
    #         print(date(*p[:3]).strftime('%Y/%m/%d'))

    # print('=========================================')

    # for p in cols:
    #     if isinstance(p,float) or isinstance(p,int):
    #         print(p)
    #     if isinstance(p,str):
    #         print(p.strip())
    #     if p.ctype == 3:
    #         print(date(*p[:3]).strftime('%Y/%m/%d'))


if __name__ == "__main__":
    read_excel()