import os

from openpyxl import load_workbook
from datetime import date

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

wb = load_workbook('MoneyControl.xlsx')
wb.iso_dates = True
ws = wb.active

MAX_ROW = ws.max_row
FORMULAS = [
    date.today().strftime('%d-%m-%Y'),
    f'=(A{MAX_ROW + 1}-A{MAX_ROW})',
    f'=ROUND((C{MAX_ROW + 1}/A{MAX_ROW})*100,0)',
    f'=ROUNDDOWN((E{MAX_ROW}+(C{MAX_ROW + 1}*0.2)),0)',
    f'=ROUNDDOWN((F{MAX_ROW}+(C{MAX_ROW + 1}*0.3)),0)',
    f'=ROUNDUP((G{MAX_ROW}+(C{MAX_ROW + 1}*0.5)),0)',
]

dataToBeAdded = [int(input()), ]
dataToBeAdded = dataToBeAdded + FORMULAS

print(dataToBeAdded)

ws.append(dataToBeAdded)

wb.save('MoneyControl.xlsx')

wb.close()


# plotting Pie Chart for the current salary breakup
wb = load_workbook('MoneyControl.xlsx')
ws = wb.active
pie = PieChart()

# selecting labels for the chart
labels = Reference(ws, min_row=1, max_row=1, min_col=5, max_col=7)

# selecting data for the chart
data = Reference(ws, min_row=ws.max_row, max_row=ws.max_row, min_col=4, max_col=7)

pie.add_data(data=data, from_rows=True, titles_from_data=True)
pie.set_categories(labels)

# Showing data labels as percentage
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True

pie.title = 'Salary Breakup'

ws.add_chart(pie, 'K2')
wb.save('MoneyControl.xlsx')
wb.close()

os.system('MoneyControl.xlsx')