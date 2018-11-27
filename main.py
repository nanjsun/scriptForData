import xlrd

data = xlrd.open_workbook('C:\Users\nanj2\Documents\lulu.xlsx')

table = data.sheets()[0]  # 通过索引顺序获取

# table = data.sheet_by_index(0)  # 通过索引顺序获取
# table = data.sheet_by_name(u'Sheet1')  # 通过名称获取

# 获取整行和整列的值（数组）
# 　　
# table.row_values(i)
print(table.nrows)

# table.col_values(i)

# 获取行数和列数
nrows = table.nrows

ncols = table.ncols

# 循环行列表数据
for i in range(nrows):
    print
table.row_values(i)

