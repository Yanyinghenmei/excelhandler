import xlwt
import random


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
    names = ['大头','飞机','钉子','老铁','象拨蚌','皮皮虾','山葵','烧柴']
    for i in range(0,len(names)):
        sheet1.write(i,0,names[i], set_style('Times New Roman',200,True))
        sheet1.write(i,1,random.randint(50,100), set_style('Times New Roman',200,True))

    wb.save('test1.xls')

if __name__ == "__main__":
    write_excel()