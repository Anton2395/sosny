# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

from openpyxl import load_workbook
import datetime
from openpyxl.styles import Border, Side
# from openpyxl.chart import ScatterChart, Reference, Series
# from openpyxl.chart.axis import DateAxis


# Load in the workbook
wb = load_workbook('./test1.xlsm')


#стиль для границ ячейки
border = Border(left=Side(border_style='thin', color='000000'),
                right=Side(border_style='thin', color='000000'),
                top=Side(border_style='thin', color='000000'),
                bottom=Side(border_style='thin', color='000000')
)

wb['Отчет по объему']['G12'].value = 1 #Наименование (тип) средства измерения
wb['Отчет по объему']['G13'].value = 1 #Номер средства измерения
wb['Отчет по объему']['G14'].value = 1 #Диапазон измерений
wb['Отчет по объему']['G15'].value = 1 #Пределы допускаемой относительной погрешности
wb['Отчет по объему']['G16'].value = 1 #Предприятие-владелец
wb['Отчет по объему']['G18'].value = datetime.date(2020, 11, 23) #Дата поверки
wb['Отчет по объему']['G19'].value = 1 #Методика поверки

#Условия поверки:
wb['Отчет по объему']['F25'].value = 1 #температура воздуха
wb['Отчет по объему']['F26'].value = 1 #относительная влажность воздуха
wb['Отчет по объему']['F27'].value = 1 #атмосферное давление



wb['Отчет по объему']['A53'].value = 1 #Ккорр
wb['Отчет по объему']['C55'].value = 1 #Поверитель



tab = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
tab1 = tab[0:13]# [0:i] - i равно значению количества столбцов -1

wb['Отчет по объему']['K30'].value = 1 #значение б(дельта)

for i in range(30, 39):
    for j in tab1:
        wb['Отчет по объему'][j+str(i)].value = i
        wb['Отчет по объему'][j+str(i)].border = border

#точки граф. "x"
wb['Лист1']['Q3'].value = 1
wb['Лист1']['Q4'].value = 3
wb['Лист1']['Q5'].value = 4
wb['Лист1']['Q6'].value = 5
#точки граф. "y"
wb['Лист1']['R3'].value = 0.6
wb['Лист1']['R4'].value = 0.11
wb['Лист1']['R5'].value = 0.14
wb['Лист1']['R6'].value = 0.16


#Создание отчёта
name = 'ot_:' + str(datetime.datetime.now()) + '.xlsx'
wb.save(name)
