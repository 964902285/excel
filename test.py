import os
import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter


file_path = 'dataImg'
wb = xlrd.open_workbook("D.xlsx")
sh1 = wb.sheet_by_name('MP1')
# 获取并打印 sheet 数量
print("sheet 数量:", wb.nsheets)

# 获取并打印 sheet 名称
print("sheet 名称:", wb.sheet_names())
# 根据 sheet 索引获取内容

# 获取并打印该 sheet 行数和列数
print(u"sheet %s 共 %d 行 %d 列" % (sh1.name, sh1.nrows, sh1.ncols))

# 获取并打印某个单元格的值
# print( "第一行第二列的值为:", sh1.cell_value(0, 1))

# 获取整行或整列的值
rows = sh1.row_values(0)  # 获取第一行内容
rows8 = sh1.row_values(7)  # 获取第一行内容
cols = sh1.col_values(2)  # 获取第二列内容

# 打印获取的行列值
# print( "第一行的值为:", rows8)
for QR_code in cols:
    if QR_code:
        print(QR_code)
# print("第二列的值为:", cols)

# 获取单元格内容的数据类型
# print( "第二行第一列的值类型为:", sh1.cell(1, 0).ctype)
sh1.write(7, 10, 1)
wb.save('test.xls')

for fileName in os.listdir(file_path):

    if len(fileName) > 21:

        QR_code = fileName[17:]
        print(QR_code)
