
__author__ = 'Daniel'

import shutil
import argparse
import xlrd
import xlwt
import re

MERGE_NAME = 'merge'
COPY_NAME = 'copy'
FILTER_NAME = 'filter'

RESULT_NAME_SUFFIX = '_result.xls'

# 主要用来测试
# python3 xlhandler.py -c copy -f test.xlsx
def copy_xl(file):
    rwb = xlrd.open_workbook(filename=file)

    wrb = xlwt.Workbook(encoding = 'utf-8')
    for sheet in rwb.sheets():
        new_sheet = wrb.add_sheet(sheet.name,cell_overwrite_ok=True)

        row_count = sheet.nrows
        col_count = sheet.ncols
        merged_cells = sheet.merged_cells
        # print(sheet.name, '\n===========\n')

        for i in range(0, row_count):
            for j in range(0, col_count):
                cell = sheet.cell(i,j)

                # sheet.new_styles(styles)

                # 判断合并的单元格
                for crange in merged_cells:
                    rlo, rhi, clo, chi = crange
                    if i == rlo and j == clo:
                        print(rlo, rhi, clo, chi)

                        # xlwt中write_merge 与 xlrd 的 merged_cells rhi/chi 相差1
                        new_sheet.write_merge(rlo,rhi-1,clo,chi-1,cell.value)
                        break 


                # 普通写操作
                new_sheet.write(i,j,cell.value)

                    

    outputName = file+RESULT_NAME_SUFFIX
    wrb.save(outputName)

# python3 xlhandler.py -c filter -k '{A}>=2' -k '{A}<=12' -f test.xlsx
# python3 xlhandler.py -c filter -k '{A}>=2' -k '{B}<=12' -f test.xlsx
# ......
def filter_xl(file, keys):

    rwb = xlrd.open_workbook(filename=file)
    wrb = xlwt.Workbook(encoding = 'utf-8')

    for sheet in rwb.sheets():
        new_sheet = wrb.add_sheet(sheet.name,cell_overwrite_ok=True)

        row_count = sheet.nrows
        col_count = sheet.ncols
        merged_cells = sheet.merged_cells

        skip = 0
        for i in range(0, row_count):

            # 处理筛选条件
            for key in keys:

                # r'[{](.*?)[}]' findall-不包含大括号
                # r'\{[^{}]*\}'  findall-包含大括号

                find_p = re.compile(r'[{](.*?)[}]', re.S)
                var_name = re.findall(find_p,key)[0]
                col = ord(var_name)-ord('A')

                # ctype :  0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
                cell = sheet.cell(i,col)
                value = cell.value
                
                replace_p = re.compile(r'\{[^{}]*\}', re.S)
                exp = re.sub(replace_p,str(value),key)

                try:
                    res = eval(exp)
                    if res:
                        # 符合筛选条件, 转录入新的Excel文件
                        for j in range(0, col_count):
                            cell = sheet.cell(i,j)

                            # 判断合并的单元格
                            for crange in merged_cells:
                                rlo, rhi, clo, chi = crange
                                if i == rlo and j == clo:
                                    print(rlo, rhi, clo, chi)

                                    # xlwt中write_merge 与 xlrd 的 merged_cells rhi/chi 相差1
                                    new_sheet.write_merge(rlo-skip,rhi-1,clo,chi-1,cell.value)
                                    break 
                            # 普通写操作
                            new_sheet.write(i-skip,j,cell.value)
                    else:
                        # 跳过
                        skip = skip+1

                except BaseException:
                    print('请检查 -k 后的表达式的操作数类型是否一致==> type:%s, value:%s, exp:%s, local:(%d,%d)' % (type(value), str(value), key, i, col))
                    return

    outputName = file+RESULT_NAME_SUFFIX
    wrb.save(outputName)
                    


# python3 xlhandler.py -c merge -k 'key>=2' -k 'key<=12' -f test1.xlsx -f test2.xlsx
def merge_xl(files, keys):
    pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument('-c', action='store', dest='command', help='excel handle command')
    parser.add_argument('-f', action='append', dest='files', default=[], help='add a excel file')
    parser.add_argument('-k', action='append', dest='keys', default=[], help='add a handle key')

    results = parser.parse_args()

    files = results.files
    keys = results.keys
    command = results.command.strip()

    # 合并
    if command == MERGE_NAME:
       pass

    # 复制
    if command == COPY_NAME:
        copy_xl(files[0])

    # 筛选
    if command == FILTER_NAME:
        filter_xl(files[0], keys)





    # fileInfo = file.split('.')
    # if len(fileInfo) < 2:
    #     raise BaseException("parameter of '-f' is wrong")

    # copyfile =  fileInfo[0] + '_copy.' + fileInfo[1]
    # # 复制备份
    # dst = shutil.copyfile(file,copyfile)
    # if dst == None:
    #     raise BaseException('copy fail')