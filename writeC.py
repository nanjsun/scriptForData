import xlrd
import xlwt


def write_c():
    c = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20181203\\C.xlsx')
    competitor = xlrd.open_workbook('C:\\Users\\nanj2\Documents\lulu\\20181203\\Competitor.xlsx')
    c_table = c.sheets()[0]
    competitor_table = competitor.sheets()[0]
    competitor_with_c = xlwt.Workbook()
    competitor_table_with_c = competitor_with_c.add_sheet('competitorWithC')
    # competitor_with_c.add_sheet(competitor_table)
    c_table_rows = c_table.nrows

    c_values = {}
    for i in range(c_table_rows):
        c_values[c_table.row_values(i)[0]] = c_table.row_values(i)[1]
        print(c_values.__sizeof__())

    competitor_table_column = int(competitor_table.ncols / 4)
    competitor_table_row = competitor_table.nrows

    for k in range(competitor_table_column * 4):
        competitor_table_with_c.write(0, k, competitor_table.row_values(0)[k])

    for i in range(competitor_table_column):
        for j in range(1, competitor_table_row):
            stock_code = competitor_table.row_values(j)[i * 4]
            if stock_code == "":
                break
            c = c_values.get(stock_code)
            competitor_table_with_c.write(j, i * 4, stock_code)
            competitor_table_with_c.write(j, i * 4 + 1, c)
    competitor_with_c.save("C:\\Users\\nanj2\Documents\lulu\\20181203\\CompetitorWithC.xls")


if __name__ == "__main__":
    write_c()


