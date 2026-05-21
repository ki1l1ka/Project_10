from logic import *
from main import *

def interface():
    print("\nДоступные действия:",
    "1 Показать массивы таблиц",
    "2 Показать таблицу",
    "3 Изменить ячейку",
    "4 Добавить столбец",
    "5 Добавить строчку",
    "6 Сумма по строке",
    "7 Сумма по столбцу",
    "8 Очистить таблицу",
    "9 Таблица 6",
    "10 Сумма по всей таблице", sep='\n')
    match int(input()):
        case 1: print([i.Show() for i in tables[input("Введите имя массива таблиц: ")].values()])
        case 2: tables[input('Таблица ')].Show()
        case 3: tables[input('Таблица ')].ChangeCell(int(input("Строчка: ")), int(input("Столбец: ")), float(input()))
        case 4: tables[input('Таблица ')].AddColumn()
        case 5: tables[input('Таблица ')].AddRow()
        case 6: print(tables[input('Таблица ')].RowSumm(int(input("Строка "))))
        case 7: print(tables[input('Таблица ')].ColumnSumm(int(input("Столбец "))))
        case 8: tables[input('Таблица ')].Clear()
        case 9: tables['t6'] = T5Create(tables['t2'], tables['t4'], Table(TableCreate(headers['t6'], lines['t6'])), int(input("год ")),  int(input("Техгология ")))
        case 10: print(tables[input('Таблица ')].SumAll())
        case 11: print(BuryReady(tables['t4'], tables['t5'], "ВВЭР-1000", "1 класс"))
while True:
    interface()
