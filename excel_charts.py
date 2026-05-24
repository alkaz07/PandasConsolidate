import openpyxl
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

excel_file_name = "test prices.xlsx"
item_categories= ['xxx', 'yyy', 'zzz']
price_categories=['A','B', 'C','D','E','F','G','H','I','K']

color_map={'A':'F00000',
           'B':'00E000', 'C':'00E000','D':'00E000',
           'E':'0000E0', 'F':'0000E0','G':'0000E0','H':'0000E0','I':'0000E0','K':'0000E0'}

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
    chart.legend=None
    chart.y_axis.delete = False
    chart.y_axis.scaling.max=100
    chart.x_axis.delete = False

    chart.varyColors=False
    # Размещение и габариты
    from_marker = AnchorMarker( col=4, row=row_min)
    to_marker = AnchorMarker(col=10, row=row_max)
    chart.anchor = TwoCellAnchor(_from=from_marker, to=to_marker)

    # Подписи к данным
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showSerName = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showLegendKey = False

    # Цвета столбиков
    for i,cat in enumerate(price_categories):
        pt = openpyxl.chart.marker.DataPoint(idx=i)
        pt.graphicalProperties.solidFill = color_map[cat]
        chart.series[0].dPt.append(pt)


    ws.add_chart(chart)
    row_min += 13
    row_max += 13

wb.save(excel_file_name)
