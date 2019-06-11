import xlwt

def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.color_index = 4
    font.height = height
    font.bold = bold
    style.font = font
    return style

def write_excel():
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet('学生', cell_overwrite_ok=True)
    row0 = ['姓名','年龄','出生日期','爱好']
    col0 = ['张三','李四','Python','小明','小红','无名']

    # 写第一行
    for i in range(0, len(row0)):
        sheet1.write(0,i,row0[i],set_style('Times New Roman', 220, True))
    
    # 写第一列
    for i in range(0, len(col0)):
        sheet1.write(i+1,0,col0[i],set_style('Times New Roman', 220, True))

    sheet1.write(1,2,'2006/12/12')
    # sheet1.write_merge(6,6,1,3,'未知')
    # sheet1.write_merge(1,2,3,3,'打游戏')
    # sheet1.write_merge(4,5,3,3,'打篮球')

    wb.save('test1.xls')

if __name__ == "__main__":
    write_excel()