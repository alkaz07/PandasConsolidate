from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

excel_file_name = "test prices.xlsx"
item_categories= ['xxx', 'yyy', 'zzz']
price_categories=['A','B', 'C','D','E','F','G','H','I','K']

color_map={'A':'red',
           'B':'green', 'C':'green','D':'green',
           'E':'blue', 'F':'blue','G':'blue','H':'blue','I':'blue','K':'blue'}

wb = load_workbook(excel_file_name)
ws = wb['Charts']


row_min = 2
row_max= 12
for category in item_categories:
    chart = BarChart()
    # Определение данных для диаграммы
    chart_values = Reference(worksheet=ws,
                             min_row=row_min+1,
                             max_row=row_max,
                             min_col=2,
                             max_col=2)
    chart_cats = Reference(ws, min_col=1, min_row=row_min + 1, max_row=row_max)
    # Добавление данных в диаграмму
    chart.add_data(chart_values)
    chart.set_categories(chart_cats)
    # Настройки внешнего вида
    chart.title="Percentage for " + category
    chart.x_axis.title='Price categories'

    ws.add_chart(chart, "D"+str(row_min))
    row_min += 14
    row_max += 14

wb.save(excel_file_name)
