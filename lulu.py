import xlrd
import xlwt

def lu() :
    data = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\0.xlsx')

    table = data.sheets()[0]  # 通过索引顺序获取

    # table = data.sheet_by_index(0)  # 通过索引顺序获取
    # table = data.sheet_by_name(u'Sheet1')  # 通过名称获取

    # 获取整行和整列的值（数组）
    # 　　
    # table.row_values(i)
    rows = table.nrows
    print(rows)

    book = xlwt.Workbook()  # 创建一个Excel
    table2 = book.add_sheet('lulu')  # 在其中创建一个名为hello的sheet
    # sheet1.write(0, 0, 'cloudox')  # 往sheet里第一行第一列写一个数据
    # sheet1.write(1, 0, 'ox')  # 往sheet里第二行第一列写一个数据
    # book.save(file)  # 创建保存文件


    j = 0

    for i in range(rows - 2) :
        if table.row_values(i)[0] == table.row_values(i + 1)[0] :
            continue


            # for k in range(4):
            #     table2.write(j,k,table.row_values(i + 1)[k])
            # j=j+1
        else:
            for k in range(4):
                table2.write(j,k,table.row_values(i)[k])
            j=j+1
    print(table.row_values(1)[0])
    book.save('new1.xls')

    # table.col_values(i)

    # # 获取行数和列数
    # nrows = table.nrows
    #
    # ncols = table.ncols

    # 循环行列表数据
    # for i in range(nrows):
    #     print
    # table.row_values(i)



if __name__ == "__main__":
    lu()