import xlrd
import xlwt
import math


def lu3():
    data = xlrd.open_workbook('11.xlsx')

    table = data.sheets()[0]  # 通过索引顺序获取

    # table = data.sheet_by_index(0)  # 通过索引顺序获取
    # table = data.sheet_by_name(u'Sheet1')  # 通过名称获取

    # 获取整行和整列的值（数组）
    # 　　
    # table.row_values(i)

    book = xlwt.Workbook()  # 创建一个Excel
    hp_index_sheet = book.add_sheet('hpIndex')  # 在其中创建一个名为hello的sheet

    rows = table.nrows
    # print(rows)
    clos = table.ncols

    first_year = {}
    for i in range(int(clos / 2)):
        current_year = table.col_values(i * 2 + 1)[0]
        hp_index_sheet.write(0, i * 2, table.col_values(i * 2)[0])
        hp_index_sheet.write(0, i * 2 + 1, table.col_values(i * 2 + 1)[0])
        for j in range(rows - 2):
            stack_id = table.col_values(i * 2)[j + 1]
            money = table.col_values(i * 2 + 1)[j + 1]
            if money == "":
                print('next col')
                break
            print("currentYear:", current_year)
            print("stackId:", stack_id)
            print("money:", money)
            if money == 0:
                hp_index_sheet.write(j + 1, i * 2, stack_id)
                hp_index_sheet.write(j + 1, i * 2 + 1, 'Wrong')
                continue

            size = math.log(float(money) / (10**0))

            if first_year.__contains__(stack_id):
                age = current_year - first_year.get(stack_id)
            else:
                age = 0
                first_year[stack_id] = current_year
            hp_index = -0.737 * size + 0.043 * size * size - 0.04 * age

            hp_index_sheet.write(j + 1, i * 2, stack_id)
            hp_index_sheet.write(j + 1, i * 2 + 1, size)
            print(age)
            print('lnSize: ', hp_index)

    book.save('lnSizeResultYuan.xls')
    # sheet1.write(0, 0, 'cloudox')  # 往sheet里第一行第一列写一个数据
    # sheet1.write(1, 0, 'ox')  # 往sheet里第二行第一列写一个数据
    # book.save(file)  # 创建保存文件
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
    lu3()
