import pandas as pd
import matplotlib.pyplot as plt
from PIL.ImtImagePlugin import field
from matplotlib import lines
from Physics import *
import json
FuelClassHeatReady = {"1 класс": 2.0, "2 класс": 0.5, "3 класс": 1.2}
# Создание таблицы
def TableExcel(FileName):
    df = pd.read_excel(f"excel/{FileName}.xlsx", header=None)
    return df.to_numpy()
def TableCreate(headers=['', ''], lines=['', '']):
    table = [[0 for i in range(len(headers))] for j in range(len(lines))]
    table[0] = headers
    for i in range(len(table)):
        table[i][0] = lines[i]
    return table
# Подсчёт объёмов РАО готовых к захоронению
def BuryReady(Table4, Tables5, FuelType, FuelClass, technology=1, StartYear=2030, EndYear=2033):
    FirstHeat = Table4.matrix[Table4.FindValue(FuelType)['row']][Table4.FindValue(FuelClass, start=(technology-1)*5)['column']].value
    HalfLife=3.0
    TotalVolume = 0
    BuryReadyVolume = 0
    Table5Row = Tables5[StartYear].FindValue(FuelType)['row']
    Table5Column = Tables5[StartYear].FindValue(FuelClass)['column']
    for year in range(StartYear, EndYear+1):
        TotalVolume += Tables5[year].matrix[Table5Row][Table5Column].value
        # print("heat", Heat(FirstHeat, year-StartYear, HalfLife))
        # print(FuelClassHeatReady[FuelClass])
        if Heat(FirstHeat, year-StartYear, HalfLife) <= FuelClassHeatReady[FuelClass]:
            BuryReadyVolume += Tables5[year].matrix[Table5Row][Table5Column].value
    # print(TotalVolume, BuryReadyVolume)
    return {"All": TotalVolume, "Ready": BuryReadyVolume}
# Классы

# Ячейка
class Cell:
    def __init__(self, value=""):
        self.value = value
    def Show(self):
        return f'self.value'
    def ChangeValue(self, value):
        self.value = value
#Таблица
class Table:
    def __init__(self, table=[['', ''], ['', '']], description='Таблица'):
        self.description = description
        self.matrix = [[Cell(table[row][column]) for column in range(len(table[row]))] for row in range(len(table))]
    def ToTable(self):
        return [[self.matrix[row][cell] for row in range(len(self.matrix))] for cell in range(len(self.matrix[0]))]
    def Show(self):
        for row in self.matrix:
            vals = [f'{x.value}' for x in row]
            print(f'{vals}')
    def Description(self, value):
        self.desсription = value
        return self.desсription
    def ChangeCell(self, row, col, value):
        self.matrix[row][col].value = value

    def GetCell(self, row, col):
        return self.matrix[row][col].value
    def RowSumm(self, row=1):
        sum = 0
        for elem in self.matrix[row][1:]:
            try:
                sum += float(elem.value)
            except (ValueError, TypeError):
                continue # Пропускаем, если это текст
        return sum
    def ColumnSumm(self, col=1):
        sum = 0
        for row_idx in range(1, len(self.matrix)):
            try:
                sum += float(self.matrix[row_idx][col].value)
            except (ValueError, TypeError):
                continue # Пропускаем, если это текст
        return sum
    def SumAll(self):
        sum = 0.0
        for row in range(1, len(self.matrix)):
            for elem in range(1, len(self.matrix[row])):
                try:
                    sum += float(self.matrix[row][elem].value)
                except (ValueError, TypeError):
                    continue # Пропускаем, если это текст
        return sum
    def AddRow(self):
        self.matrix.append([Cell() for i in range(len(self.matrix[0]))])
    def AddColumn(self, header="head"):
        for i in range(len(self.matrix)):
            self.matrix[i].append(Cell())
        self.matrix[0][-1].value = header
    def DeleteRow(self):
        self.matrix.pop(-1)
    def DeleteColumn(self):
        for i in range(len(self.matrix[0])):
            self.matrix[i].pop(-1)
    def FillRow(self, Row, RowValues):
        for i in range(len(self.matrix[Row])):
            self.matrix[Row][i].value = RowValues[i]

    def FillColumn(self, Column, ColumnValues):
        for i in range(len(ColumnValues)):
            self.matrix[i][Column].value = ColumnValues[i]
    def FindValue(self, Value, start=0):
        target = str(Value).strip().lower()
        for row in range(len(self.matrix)):
            for elem in range(start, len(self.matrix[row])):
                current = str(self.matrix[row][elem].value).strip().lower()
                if current == target:
                    return {'row': row, 'column': elem}
        print(f"Ошибка: Значение '{Value}' не найдено в таблице!")
        return {'row': 0, 'column': 0}

    def Clear(self):
        self.matrix = [[Cell() for i in range(len(self.matrix[0]))] for j in range(len(self.matrix))]

class CreatedTable(Table):
    def T5Create(self, t2, t3, t5, year, technology=1):
        if technology == 1:
            start = 0
        else:
            start = 5
        for i in range(1, len(t5.matrix)):
            fuel = t5.matrix[i][0].value
            for j in range(1, len(t5.matrix[i])):
                FuelClass = t5.matrix[0][j].value
                t2Value = t2.matrix[t2.FindValue(year)['row']][t2.FindValue(fuel)['column']].value
                t3Value = t3.matrix[t3.FindValue(fuel)['row']][t3.FindValue(FuelClass, start)['column']].value
                t5.ChangeCell(i, j, round(t2Value * t3Value, 1))
        return t5

    def T6Create(self, table6, tables5, table4, year=2030, technology=1):
        Start = 5 * (technology - 1)
        table5 = tables5[year]
        for i in range(1, len(table6.matrix[0])):
            WholeHeat = 0
            Divider = 0
            FuelClass = table6.matrix[0][i].value
            for j in range(1, len(table5.matrix)):
                FuelType = table5.matrix[j][0].value
                T5Row = table5.FindValue(FuelType)['row']
                T5Column = table5.FindValue(FuelClass)['column']
                T4Row = table4.FindValue(FuelType)['row']
                T4Column = table4.FindValue(FuelClass, start=Start)['column']
                Elem1 = table5.matrix[T5Row][T5Column].value
                Elem2 = table4.matrix[T4Row][T4Column].value
                WholeHeat += (Elem1 * Elem2)
                Divider += Elem1
            AverageHeat = round(WholeHeat / Divider, 3)
            T6Column = table6.FindValue(FuelClass)['column']
            table6.matrix[1][T6Column].value = round(AverageHeat, 1)
        return table6

    def T7Create(self, Table4, Tables5, table7, technology=1, startyear=2030, endyear=2050):
        # Список всех видов топлива для суммирования
        reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']
        ColClass1 = 1
        ColClass2 = 3
        ColClass3 = 5
        FuelClasses = ["1 класс", "2 класс", "3 класс"]
        table7.DeleteRow()
        for i in range(startyear, endyear + 1):
            table7.AddRow()
            table7.matrix[-1][0].value = i

            # Для каждого класса считаем сумму
            for idx, fuel_class in enumerate(FuelClasses):
                total_all = 0
                total_ready = 0

                for fuel in reactors:
                    res = BuryReady(Table4, Tables5, fuel, fuel_class, technology, StartYear=startyear, EndYear=i)
                    total_all += res["All"]
                    total_ready += res["Ready"]

                col_idx = 1 + (idx * 2)  # 1, 3, 5
                table7.matrix[-1][col_idx].value = round(total_all, 1)
                table7.matrix[-1][col_idx + 1].value = round(total_ready, 1)

        return table7

    def RunSensitivityAnalysis(self, base_t4, tables_t5_dict, table_t2, table_t3, technology=1, start_year=2030,
                               end_year=2050, min_pct=50, max_pct=150):
        import copy
        coefficients = [x / 100.0 for x in range(min_pct, max_pct + 1, 5)]
        ready_volumes_2050 = []

        all_reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']
        headers = {
            't5': ['Завод', '1 класс', '2 класс', '3 класс', '4 класс'],
            't7': ["Год", '1 класс', "Готово к захоронению", '2 класс', "Готово к захоронению", '3 класс',
                   "Готово к захоронению"]
        }
        lines = {'t5': ['Завод'] + all_reactors}

        start_col_shift = 5 * (technology - 1)

        for coeff in coefficients:
            temp_t4 = copy.deepcopy(base_t4)

            for reactor in all_reactors:
                row_idx = temp_t4.FindValue(reactor)['row']
                col_idx = temp_t4.FindValue("1 класс", start=start_col_shift)['column']
                original_value = float(temp_t4.matrix[row_idx][col_idx].value)
                temp_t4.matrix[row_idx][col_idx].value = original_value * coeff

            temp_t5_series = {}
            for year in range(start_year, end_year + 1):
                temp_t5_series[year] = CreatedTable(TableCreate(headers['t5'], lines['t5']))
                temp_t5_series[year].T5Create(t2=table_t2, t3=table_t3, t5=temp_t5_series[year], year=year,
                                              technology=technology)

            temp_t7 = CreatedTable(TableCreate(headers['t7']))
            temp_t7.T7Create(temp_t4, temp_t5_series, temp_t7, technology=technology, startyear=start_year,
                             endyear=end_year)

            # строка для конечного года расчета (2050)
            row_end_year_idx = temp_t7.FindValue(end_year)['row']
            if row_end_year_idx == 0:
                # Предохранитель
                row_end_year_idx = len(temp_t7.matrix) - 1

            vol_ready_end_year = float(temp_t7.matrix[row_end_year_idx][2].value)
            ready_volumes_2050.append(round(vol_ready_end_year, 1))

        percentages = [int(c * 100) for c in coefficients]
        return percentages, ready_volumes_2050


class Graph:
    def __init__(self, table=None, column=False, row=False):
        self.graph = dict()
        if table:
            # Строки от колонок и наоборот
            if column:
                for i in range(1, len(table.matrix)):
                    self.graph[table.matrix[i][0].value] = table.matrix[i][column].value
            elif row:
                for i in range(1, len(table.matrix[0])):
                    self.graph[table.matrix[0][i].value] = table.matrix[row][i].value

    @classmethod
    def FromTableSeries(cls, table_dict, fuel_type, column_name):
        instance = cls()
        for year, table in table_dict.items():
            # строка с реактором
            res = table.FindValue(fuel_type)
            # колонка с классом
            col_res = table.FindValue(column_name)

            if res and col_res:
                val = table.matrix[res['row']][col_res['column']].value
                try:
                    instance.graph[year] = float(val)
                except (ValueError, TypeError):
                    continue  # Если попал заголовок, просто пропускаем
        return instance


import matplotlib.pyplot as plt
import os
from matplotlib.ticker import MultipleLocator

class NuclearDataVisualizer:
    def __init__(self):
        self.datasets = {}

    def add_graph(self, label, graph_obj):

        # данные из словаря внутри объекта Graph
        raw_data = graph_obj.graph

        # Сортируем по годам
        sorted_keys = sorted(raw_data.keys(), key=lambda x: float(x))
        x_values = [float(k) for k in sorted_keys]
        y_values = [float(raw_data[k]) for k in sorted_keys]

        self.datasets[label] = (x_values, y_values)

    def add_plot_to_doc(vis_obj, title):
        if not vis_obj.datasets:
            return
        buf = io.BytesIO()
        plt.figure(figsize=(10, 5))

        for label, (x, y) in vis_obj.datasets.items():
            plt.plot(x, y, marker='o', label=label)

        plt.title(title)
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
        plt.grid(True, linestyle='--', alpha=0.6)

        ax = plt.gca()
        ax.xaxis.set_major_locator(MultipleLocator(5))

        plt.tight_layout()
        plt.savefig(buf, format='png', bbox_inches='tight')
        doc.add_picture(buf, width=Inches(5.5))
        plt.close()
    def clear(self):
        self.datasets = {}


class TableVisualizer:
    def __init__(self, table_obj):
        self.table = table_obj
        self.description = getattr(table_obj, 'description', 'Таблица данных')

    def plot_table(self, exclude_rows=None, exclude_cols=None, title=None, save_path=None):

        # Превращаем матрицу объектов Cell в список списков
        raw_data = [[cell.value for cell in row] for row in self.table.matrix]

        # Фильтрация
        if exclude_rows:
            raw_data = [row for i, row in enumerate(raw_data) if i not in exclude_rows]

        if exclude_cols:
            filtered_data = []
            for row in raw_data:
                filtered_row = [val for j, val in enumerate(row) if j not in exclude_cols]
                filtered_data.append(filtered_row)
            raw_data = filtered_data

        if not raw_data:
            print("Ошибка: После фильтрации таблица пуста!")
            return

        # Настройка отображения
        fig, ax = plt.subplots(figsize=(12, len(raw_data) * 0.6))
        ax.axis('off')

        # Создание таблицы matplotlib
        thead = raw_data[0]
        tbody = raw_data[1:]

        table_plot = ax.table(cellText=tbody, colLabels=thead,
                              loc='center', cellLoc='center')

        table_plot.auto_set_font_size(False)
        table_plot.set_fontsize(10)
        table_plot.scale(1.2, 1.8)  # Масштаб ячеек по высоте и ширине

        # Оформление заголовков
        for (row, col), cell in table_plot.get_celld().items():
            if row == 0:
                cell.set_text_props(weight='bold', color='white')
                cell.set_facecolor('#40466e')

        plt.title(title or self.description, fontsize=14, pad=20)

        if save_path:
            plt.savefig(save_path, bbox_inches='tight', dpi=300)
        plt.show()