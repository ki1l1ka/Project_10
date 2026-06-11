from matplotlib.ticker import MultipleLocator
import matplotlib.pyplot as plt
from PIL.ImtImagePlugin import field
import copy
import os
import pandas as pd
from Physics import *
import streamlit as st
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
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
    HalfLife= st.session_state.HalfLife[FuelClass]
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
        # rows = len(table)
        # cols = len(table[0])
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
                continue
        return sum
    def SumAll(self):
        sum = 0.0
        for row in range(1, len(self.matrix)):
            for elem in range(1, len(self.matrix[row])):
                try:
                    sum += float(self.matrix[row][elem].value)
                except (ValueError, TypeError):
                    continue
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

    def UpdateFromDataFrame(self, old_df, new_df):
        if old_df is None or new_df is None:
            return False
        # Обходим строки и колонки
        for r in range(len(new_df)):
            for c in range(len(new_df.columns)):
                old_val = old_df.iloc[r, c]
                new_val = new_df.iloc[r, c]
                if old_val != new_val:
                    # Приведение типов данных
                    if pd.isna(new_val) or str(new_val).strip() == "":
                        typed_val = None
                    else:
                        try:
                            typed_val = float(new_val) if '.' in str(new_val) else int(new_val)
                        except ValueError:
                            typed_val = new_val
                    if (r + 1) < len(self.matrix) and c < len(self.matrix[r + 1]):
                        self.matrix[r + 1][c].value = typed_val
        return True

    def ToDataFrame(self):
        raw_matrix = [[cell.value for cell in row] for row in self.matrix]
        if len(raw_matrix) == 0:
            return pd.DataFrame()
        max_cols = max(len(row) for row in raw_matrix)
        headers = []
        for cell_val in raw_matrix[0]:
            if cell_val is None:
                val_str = ""
            elif isinstance(cell_val, (list, tuple)):
                val_str = " ".join([str(v).strip() for v in cell_val if v is not None])
            else:
                val_str = str(cell_val).strip()
            if val_str.lower() in ["nan", "none"]: val_str = ""
            headers.append(val_str)

        if len(headers) < max_cols:
            headers.extend([""] * (max_cols - len(headers)))
        else:
            headers = headers[:max_cols]

        # уникальность имен
        seen = {}
        for idx, h in enumerate(headers):
            if h in seen:
                seen[h] += 1
                headers[idx] = h + (" " * seen[h])
            else:
                seen[h] = 0
        body_data = []
        for row in raw_matrix[1:]:
            if len(row) == 0: continue
            new_row = []
            for cell_val in row:
                if cell_val is None or str(cell_val).lower() in ["nan", "none"]:
                    new_row.append("")
                else:
                    new_row.append(round(cell_val, 2) if isinstance(cell_val, (int, float)) else cell_val)
            if len(new_row) < max_cols:
                new_row.extend([""] * (max_cols - len(new_row)))
            else:
                new_row = new_row[:max_cols]
            body_data.append(new_row)
        return pd.DataFrame(body_data, columns=headers)

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
            res = table.FindValue(fuel_type)
            col_res = table.FindValue(column_name)

            if res and col_res:
                val = table.matrix[res['row']][col_res['column']].value
                try:
                    instance.graph[year] = float(val)
                except (ValueError, TypeError):
                    continue  # Если попал заголовок - пропускаем
        return instance

class ProcessingPlant:

    def __init__(self, name="Завод №1", folder_path="excel/Завод 1"):
        self.name = name
        self.folder_path = folder_path

        # Исходные таблицы
        self.t1 = None
        self.t2 = None
        self.t3 = None
        self.t4 = None

        # Расчетные таблицы (словари по технологиям 1, 2
        self.t5_series = {1: {}, 2: {}}
        self.t6_series = {1: {}, 2: {}}
        self.t7 = {1: None, 2: None}

        self.LoadPlantData()

    def LoadPlantData(self):
        import os
        import pandas as pd
        from logic import Table
        global_excel_root = "excel"

        # t1 считывается из корня
        t1_global_path = os.path.join(global_excel_root, "t1.xlsx")
        if not os.path.exists(t1_global_path):
            raise FileNotFoundError(f"Критическая ошибка: Общая таблица Т1 не найдена по пути '{t1_global_path}'!")

        self.t1 = Table(pd.read_excel(t1_global_path, header=None).to_numpy(),
                        description=f"Таблица Т1. Общие исходные данные по образованию ОЯТ (тонны тяжелого металла)")

        # Проверяем наличие файлов (t2, t3, t4) в подпапке конкретного завода
        if not os.path.exists(self.folder_path):
            return

        for t_name in ['t2', 't3', 't4']:
            file_p = os.path.join(self.folder_path, f"{t_name}.xlsx")
            if not os.path.exists(file_p):
                raise FileNotFoundError(
                    f"В папке завода {self.folder_path} отсутствует обязательный файл {t_name}.xlsx")

        try:
            # Считываем таблицы завода (t2, t3, t4) из его собственной папки
            self.t2 = Table(pd.read_excel(os.path.join(self.folder_path, "t2.xlsx"), header=None).to_numpy(),
                            description=f"Таблица Т2. Данные по загрузке завода ОЯТ — {self.name}")
            self.t3 = Table(pd.read_excel(os.path.join(self.folder_path, "t3.xlsx"), header=None).to_numpy(),
                            description=f"Таблица Т3. Образование РАО различных классов — {self.name}")
            self.t4 = Table(pd.read_excel(os.path.join(self.folder_path, "t4.xlsx"), header=None).to_numpy(),
                            description=f"Таблица Т4. Тепловыделение РАО сразу после переработки — {self.name}")

            print(f"Данные для '{self.name}' успешно инициализированы (Т1 взята из общего корня)")
        except Exception as e:
            raise IOError(f"Критическая ошибка при парсинге Excel-таблиц для {self.name}: {e}")

    def CalculateMatrices(self, technology, start_year, end_year):
        self.t5_series[technology] = {}
        self.t6_series[technology] = {}
        self.t7[technology] = None

        all_reactors = ["ВВЭР-1000", "ВВЭР-440", "БН-600", "БН-800", "РБМК"]
        headers = {
            "t5": ["Завод", "1 класс", "2 класс", "3 класс", "4 класс"],
            "t6": ["Завод", "1 класс", "2 класс", "3 класс", "4 класс"],
            "t7": ["Год",
                "1 класс", "Готово к захоронению",
                "2 класс", "Готово к захоронению",
                "3 класс", "Готово к захоронению"],
        }

        lines = {"t5": ["Завод"] + all_reactors, "t6": ["Завод", "..."]}
        # Считаем Т5 и Т6 по годам
        for y in range(start_year, end_year + 1):
            self.t5_series[technology][y] = CreatedTable(
                TableCreate(headers["t5"], lines["t5"])
            )
            self.t5_series[technology][y].T5Create(
                t2=self.t2,
                t3=self.t3,
                t5=self.t5_series[technology][y],
                year=y,
                technology=technology,
            )
            self.t6_series[technology][y] = CreatedTable(
                TableCreate(headers["t6"], lines["t6"])
            )
            self.t6_series[technology][y].T6Create(
                table6=self.t6_series[technology][y],
                tables5=self.t5_series[technology],
                table4=self.t4,
                year=y,
                technology=technology,
            )
        # Расчет итоговой Т7 для выбранной технологии (Внутри logic.py)
        self.t7[technology] = CreatedTable(
            TableCreate(headers["t7"]),
            description=f"Количество РАО накопительным итогом (Технология {technology}) — {self.name}",
        )
        self.t7[technology].T7Create(
            self.t4,
            self.t5_series[technology],
            self.t7[technology],
            technology=technology,
            startyear=start_year,
            endyear=end_year
        )

    def RunSensitivityAnalysis(self, technology, start_year, end_year, min_pct, max_pct):
        coefficients = [x / 100.0 for x in range(min_pct, max_pct + 1, 5)]
        ready_volumes = []
        all_reactors = ["ВВЭР-1000", "ВВЭР-440", "БН-600", "БН-800", "РБМК"]
        headers = {
            "t5": ["Завод", "1 класс", "2 класс", "3 класс", "4 класс"],
            "t7": [
                "Год",
                "1 класс",
                "Готово к захоронению",
                "2 класс",
                "Готово к захоронению",
                "3 класс",
                "Готово к захоронению"
            ],
        }
        lines = {"t5": ["Завод"] + all_reactors}
        start_col_shift = 5 * (technology - 1)
        for coeff in coefficients:
            temp_t4 = copy.deepcopy(self.t4)
            for reactor in all_reactors:
                row_idx = temp_t4.FindValue(reactor)["row"]
                col_idx = temp_t4.FindValue("1 класс", start=start_col_shift)[
                    "column"
                ]
                original_value = float(temp_t4.matrix[row_idx][col_idx].value)
                temp_t4.matrix[row_idx][col_idx].value = original_value * coeff
            temp_t5_series = {}
            for year in range(start_year, end_year + 1):
                temp_t5_series[year] = CreatedTable(
                    TableCreate(headers["t5"], lines["t5"])
                )
                temp_t5_series[year].T5Create(
                    t2=self.t2,
                    t3=self.t3,
                    t5=temp_t5_series[year],
                    year=year,
                    technology=technology,
                )
            temp_t7 = CreatedTable(TableCreate(headers["t7"]))
            temp_t7.T7Create(
                temp_t4,
                temp_t5_series,
                temp_t7,
                technology=technology,
                startyear=start_year,
                endyear=end_year,
            )
            row_end_idx = temp_t7.FindValue(end_year)["row"]
            if row_end_idx == 0:
                row_end_idx = len(temp_t7.matrix) - 1
            vol_ready = float(temp_t7.matrix[row_end_idx][2].value)
            ready_volumes.append(round(vol_ready, 1))
        percentages = [int(c * 100) for c in coefficients]
        return percentages, ready_volumes
class BurialSite:
    def __init__(self, ClassA="1 класс", ClassACapacity=0, ClassB="2 класс", ClassBCapacity=0, ClassAMaxTempo=1000, ClassBMaxTempo=1500):
        self.Capacities = {str(ClassA): float(ClassACapacity), str(ClassB): float(ClassBCapacity)}
        self.FullCapacity = ClassACapacity + ClassBCapacity # Предполагается захоронение только двух классов на объект
        self.Full = False
        if abs(int(ClassA.strip()[0]) - int(ClassB.strip()[0])) > 1:
            print('''Возможная ошибка! Убедитесь в верности заданных классов.
                    Для глубинного - 1,2 классы.
                    Для приповерхностного - 3,4 классы.''')
    def Load(self, wasteClass, wasteAmount):
        if wasteClass not in self.Capacities.keys():
            print(f"Ошибка: Класс '{wasteClass}' не поддерживается данным объектом.")
            return
        if self.Capacities[wasteClass] >= wasteAmount:
            self.Capacities[wasteClass] -= wasteAmount
            self.FullCapacity -= wasteAmount
            # print("Загрузка прошла успешно")
        else:
            if (self.Capacities[wasteClass] - wasteAmount) < 0:
                self.Capacities[wasteClass] = 0
                left = abs(self.Capacities[wasteClass] - wasteAmount)
                print(f"Удалось загрузить {wasteAmount-left}")
                print(f"Осталось {left}")
                return left
            # print("Хранилище полностью заполнено.")
            self.Full = True
    def IsFull(self):
        return self.Full

class NuclearDataVisualizer:
    def __init__(self):
        self.datasets = {}

    def add_graph(self, label, graph_obj):
        raw_data = graph_obj.graph
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

    def GetT7PlotData(self, active_t7_obj):
        plot_data = {
            "1 класс (Всего)": [], "1 класс (Готово)": [],
            "2 класс (Всего)": [], "2 класс (Готово)": [],
            "3 класс (Всего)": [], "3 класс (Готово)": []
        }# Сейчас не учитывается 4-й класс

        if active_t7_obj is None or not hasattr(active_t7_obj, 'matrix'):
            return plot_data
        config = [
            (1, 2, "1 класс"),
            (3, 4, "2 класс"),
            (5, 6, "3 класс")
        ]
        for col_all, col_ready, class_key in config:
            for r in range(1, len(active_t7_obj.matrix)):
                # Всего
                val_all = active_t7_obj.matrix[r][col_all].value
                num_all = float(val_all) if val_all is not None and str(val_all).strip() != "" else 0.0
                plot_data[f"{class_key} (Всего)"].append(num_all)

                # Готово
                val_ready = active_t7_obj.matrix[r][col_ready].value
                num_ready = float(val_ready) if val_ready is not None and str(val_ready).strip() != "" else 0.0
                plot_data[f"{class_key} (Готово)"].append(num_ready)

        return plot_data

    def clear(self):
        self.datasets = {}

    def CreateT7PlotInCore(self, active_t7_obj, years_range):

        fig, ax = plt.subplots(figsize=(9, 4.2))
        t7_data = self.GetT7PlotData(active_t7_obj)

        colors = {
            "1 класс (Всего)": "#d9534f", "1 класс (Готово)": "#942a27",
            "2 класс (Всего)": "#f0ad4e", "2 класс (Готово)": "#b57d28",
            "3 класс (Всего)": "#5cb85c", "3 класс (Готово)": "#2b702b"
        }

        for label_name, y_values in t7_data.items():
            if len(years_range) == len(y_values) and label_name in colors:
                line_style = '--' if 'Всего' in label_name else '-'
                marker_style = 'o' if 'Всего' in label_name else 's'
                marker_size = 3 if 'Всего' in label_name else 4
                line_width = 1.5 if 'Всего' in label_name else 2.5

                ax.plot(years_range, y_values, linestyle=line_style, marker=marker_style,
                        markersize=marker_size, color=colors[label_name],
                        label=label_name, linewidth=line_width)

        ax.set_xlabel("Годы", fontsize=10)
        ax.set_ylabel("Объем РАО, м³", fontsize=10)
        ax.grid(True, linestyle='--', alpha=0.5)
        ax.xaxis.set_major_locator(MultipleLocator(5))
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)
        plt.tight_layout()
        return fig

    def CreateBurialPlotsInCore(self, plants_session_dict, start_yr, end_yr):
        burial_plot_data = GetBurialAccumulationData(plants_session_dict, start_yr, end_yr)
        yrs = burial_plot_data["years"]
        # ПГЗР
        fig_pgzr, ax_pgzr = plt.subplots(figsize=(9, 3.8))
        pgzr_config = [
            ("1 класс (ПГЗР)", "#942a27", 40000.0, "1 класс (Лимит 40к м³)"),
            ("2 класс (ПГЗР)", "#b57d28", 60000.0, "2 класс (Лимит 60к м³)")
        ]
        for key, color, limit, label_text in pgzr_config:
            ax_pgzr.plot(yrs, burial_plot_data[key], linestyle='-', marker='o', markersize=4, color=color,
                         label=label_text, linewidth=2.5)
            ax_pgzr.axhline(y=limit, color=color, linestyle='--', alpha=0.6, linewidth=1.2)
        ax_pgzr.set_xlabel("Годы", fontsize=9)
        ax_pgzr.set_ylabel("Заполнение, м³", fontsize=9)
        ax_pgzr.grid(True, linestyle='--', alpha=0.4)
        ax_pgzr.xaxis.set_major_locator(MultipleLocator(5))
        ax_pgzr.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)
        plt.tight_layout()
        # ППЗР
        fig_ppzr, ax_ppzr = plt.subplots(figsize=(9, 3.8))
        ppzr_config = [
            ("3 класс (ППЗР)", "#2b702b", 40000.0, "3 класс (Лимит 40к м³)"),
            ("4 класс (ППЗР)", "#2a6496", 100000.0, "4 класс (Лимит 100к м³)")
        ]
        for key, color, limit, label_text in ppzr_config:
            ax_ppzr.plot(yrs, burial_plot_data[key], linestyle='-', marker='s', markersize=4, color=color,
                         label=label_text, linewidth=2.5)
            ax_ppzr.axhline(y=limit, color=color, linestyle='--', alpha=0.6, linewidth=1.2)
        ax_ppzr.set_xlabel("Годы", fontsize=9)
        ax_ppzr.set_ylabel("Заполнение, м³", fontsize=9)
        ax_ppzr.grid(True, linestyle='--', alpha=0.4)
        ax_ppzr.xaxis.set_major_locator(MultipleLocator(5))
        ax_ppzr.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)
        plt.tight_layout()

        return fig_pgzr, fig_ppzr

    def CreateReactorDefaultPlot(self, selected_reactors, years_range):
        fig, ax = plt.subplots(figsize=(9, 4.0))
        for reactor in selected_reactors:
            volumes = np.cumsum(np.random.randint(15, 60, len(years_range)))
            ax.plot(years_range, volumes, marker='o', label=reactor, linewidth=2)

        ax.set_xlabel("Годы", fontsize=10)
        ax.set_ylabel("Значение показателей", fontsize=10)
        ax.grid(True, linestyle='--', alpha=0.5)
        ax.xaxis.set_major_locator(MultipleLocator(5))
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)
        plt.tight_layout()
        return fig


class TableVisualizer:
    def __init__(self, table_obj):
        self.table = table_obj
        self.description = getattr(table_obj, 'description', 'Таблица данных')

    def plot_table(self, exclude_rows=None, exclude_cols=None, title=None, save_path=None):

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
        fig, ax = plt.subplots(figsize=(12, len(raw_data) * 0.6))
        ax.axis('off')
        thead = raw_data[0]
        tbody = raw_data[1:]

        table_plot = ax.table(cellText=tbody, colLabels=thead,
                              loc='center', cellLoc='center')

        table_plot.auto_set_font_size(False)
        table_plot.set_fontsize(10)
        table_plot.scale(1.2, 1.8)

        # заголовки
        for (row, col), cell in table_plot.get_celld().items():
            if row == 0:
                cell.set_text_props(weight='bold', color='white')
                cell.set_facecolor('#40466e')

        plt.title(title or self.description, fontsize=14, pad=20)

        if save_path:
            plt.savefig(save_path, bbox_inches='tight', dpi=300)
        plt.show()
def ExportTableToExcel(table_obj, label_name="таблица"):
    import os
    import pandas as pd

    if table_obj is None or not hasattr(table_obj, 'matrix'):
        return False

    folder = "отчёты"
    if not os.path.exists(folder):
        os.makedirs(folder)
    clean_label = label_name.replace(" ", "_").replace("(", "").replace(")", "")
    n = 1
    while os.path.exists(f"{folder}/экспорт_{clean_label}_{n}.xlsx"):
        n += 1
    filename = f"{folder}/экспорт_{clean_label}_{n}.xlsx"

    raw_matrix = [[cell.value for cell in row] for row in table_obj.matrix]
    if len(raw_matrix) == 0:
        return False

    headers = []
    for cell_val in raw_matrix[0]:
        val_str = str(cell_val).strip() if pd.notna(cell_val) else ""
        if val_str.lower() == "nan" or val_str == "None":
            val_str = ""
        headers.append(val_str)

    max_cols = max(len(row) for row in raw_matrix)
    if len(headers) < max_cols:
        headers.extend([""] * (max_cols - len(headers)))
    else:
        headers = headers[:max_cols]

    seen = {}
    for idx, h in enumerate(headers):
        if h in seen:
            seen[h] += 1
            headers[idx] = h + (" " * seen[h])
        else:
            seen[h] = 0

    body_data = []
    for row in raw_matrix[1:]:
        new_row = []
        for cell_val in row:
            if pd.isna(cell_val) or str(cell_val).lower() == "nan" or str(cell_val) == "None":
                new_row.append("")
            else:
                new_row.append(cell_val)
        if len(new_row) < max_cols:
            new_row.extend([""] * (max_cols - len(new_row)))
        else:
            new_row = new_row[:max_cols]
        body_data.append(new_row)

    try:
        df_export = pd.DataFrame(body_data, columns=headers)
        df_export.to_excel(filename, index=False)
        return filename
    except Exception as e:
        print(f"Ошибка сохранения Excel: {e}")
        return False


class ReportConstructor:
    def __init__(self):
        self.SelectedElements = []

    def GetAvailableElements(self, plants_session_dict, start_yr, end_yr):
        registry = []
        first_plant_id = list(plants_session_dict.keys())[0] if plants_session_dict else None
        # Т1
        registry.append({
            "key": "global_t1",
            "type": "table",
            "group": "Глобальные данные",
            "label": "Таблица Т1. Исходные данные по образованию ОЯТ по годам",
            "plant_name": None,
            "field_name": "t1",
            "year": None
        })
        for p_id, p_info in plants_session_dict.items():
            p_name = p_info["name"]
            group_label = f"Завод: {p_name}"
            registry.append({"key": f"t2_{p_id}", "type": "table", "group": group_label,
                             "label": "Таблица Т2. Данные по загрузке завода ОЯТ", "plant_name": p_name,
                             "field_name": "t2", "year": None})
            registry.append({"key": f"t3_{p_id}", "type": "table", "group": group_label,
                             "label": "Таблица Т3. Нормативы образования РАО", "plant_name": p_name, "field_name": "t3",
                             "year": None})
            registry.append({"key": f"t4_{p_id}", "type": "table", "group": group_label,
                             "label": "Таблица Т4. Удельное тепловыделение РАО", "plant_name": p_name,
                             "field_name": "t4", "year": None})

            # Расчетные Т5 и Т6
            for tech in [1, 2]:
                for yr in range(start_yr, end_yr + 1):
                    registry.append({
                        "key": f"t5_{p_id}_{tech}_{yr}", "type": "table", "group": group_label,
                        "label": f"Таблица Т5. Объемы РАО (Технология {tech}, Год {yr})",
                        "plant_name": p_name, "field_name": "t5_series", "tech": tech, "year": yr
                    })
                    registry.append({
                        "key": f"t6_{p_id}_{tech}_{yr}", "type": "table", "group": group_label,
                        "label": f"Таблица Т6. Среднее тепловыделение РАО (Технология {tech}, Год {yr})",
                        "plant_name": p_name, "field_name": "t6_series", "tech": tech, "year": yr
                    })

                # Итоговая накопительная Т7
                registry.append({
                    "key": f"t7_{p_id}_{tech}", "type": "table", "group": group_label,
                    "label": f"Таблица Т7. Итоговая накопительная ведомость РАО (Технология {tech})",
                    "plant_name": p_name, "field_name": "t7", "tech": tech, "year": None
                })

            # Графики по текущему заводу
            registry.append({"key": f"g_t7_{p_id}", "type": "graph", "group": group_label,
                             "label": "График Т7. Накопление и готовность РАО 1 класса", "plant_name": p_name,
                             "field_name": "graph_t7"})
            registry.append({"key": f"g_sens_{p_id}", "type": "graph", "group": group_label,
                             "label": "График чувствительности. Зависимость объемов от тепловыделения",
                             "plant_name": p_name, "field_name": "graph_sens"})

        return registry


def SimulateAllBurials(plants_session_dict, start_yr, end_yr):
    BurialSiteDeep = BurialSite("1 класс", 40000, "2 класс", 60000)
    BurialSiteSurface = BurialSite("3 класс", 40000, "4 класс", 100000)
    deficit_records = []
    for yr in range(start_yr, end_yr + 1):
        year_ready_1cl = 0.0
        year_ready_2cl = 0.0
        year_ready_3cl = 0.0
        year_ready_4cl = 0.0


        for p_id, p_info in plants_session_dict.items():
            plant_key = f"plant_object_{p_id}"

            if plant_key in st.session_state and st.session_state[plant_key] is not None:
                plant_obj = st.session_state[plant_key]
                tech = p_info["technology"]
                t7_dict = getattr(plant_obj, "t7", {})
                t7_obj = t7_dict.get(tech) if isinstance(t7_dict, dict) else t7_field

                if t7_obj is not None and hasattr(t7_obj, 'matrix'):
                    row_idx = 0
                    for r in range(1, len(t7_obj.matrix)):
                        if t7_obj.matrix[r].value == yr:
                            row_idx = r
                            break

                    # суммируем кубометры
                    if t7_obj.matrix[r][0].value == yr:
                        matrix_row = t7_obj.matrix[row_idx]
                        year_ready_1cl += float(matrix_row[2].value or 0.0)
                        year_ready_2cl += float(matrix_row[4].value or 0.0)
                        year_ready_3cl += float(matrix_row[6].value or 0.0)
                        if len(matrix_row) > 8: year_ready_4cl += float(matrix_row[8].value or 0.0)

        # Распределяем суммарные объемы РАО отрасли по объектам захоронения
        left_1 = BurialSiteDeep.Load("1 класс", year_ready_1cl)
        left_2 = BurialSiteDeep.Load("2 класс", year_ready_2cl)
        left_3 = BurialSiteSurface.Load("3 класс", year_ready_3cl)
        left_4 = BurialSiteSurface.Load("4 класс", year_ready_4cl)

        if (left_1 or 0) > 0 or (left_2 or 0) > 0 or (left_3 or 0) > 0 or (left_4 or 0) > 0:
            deficit_records.append({
                "Год": yr,
                "1 класс (деф, м³)": round(left_1 or 0.0, 1),
                "2 класс (деф, м³)": round(left_2 or 0.0, 1),
                "3 класс (деф, м³)": round(left_3 or 0.0, 1),
                "4 класс (деф, м³)": round(left_4 or 0.0, 1)
            })

    return deficit_records

def GetBurialAccumulationData(plants_session_dict, start_yr, end_yr):
    # Инициализируем чистые хранилища
    BurialSiteDeep = BurialSite("1 класс", 40000, "2 класс", 60000)
    BurialSiteSurface = BurialSite("3 класс", 40000, "4 класс", 100000)
    accumulation_data = {
        "years": [],
        "1 класс (ПГЗР)": [], "2 класс (ПГЗР)": [],
        "3 класс (ППЗР)": [], "4 класс (ППЗР)": []
    }

    # Счетчики объемов в захоронениях
    filled_1cl = 0.0
    filled_2cl = 0.0
    filled_3cl = 0.0
    filled_4cl = 0.0

    for yr in range(start_yr, end_yr + 1):
        year_ready_1cl = 0.0
        year_ready_2cl = 0.0
        year_ready_3cl = 0.0
        year_ready_4cl = 0.0

        # Собираем готовые объемы РАО
        for p_id, p_info in plants_session_dict.items():
            import streamlit as st
            plant_key = f"plant_object_{p_id}"

            if plant_key in st.session_state and st.session_state[plant_key] is not None:
                plant_obj = st.session_state[plant_key]
                tech = p_info["technology"]
                t7_dict = getattr(plant_obj, "t7", {})
                t7_obj = t7_dict.get(tech) if isinstance(t7_dict, dict) else None

                if t7_obj is not None and hasattr(t7_obj, 'matrix'):
                    row_idx = 0
                    for r in range(1, len(t7_obj.matrix)):
                        if t7_obj.matrix[r][0].value == yr:
                            row_idx = r
                            break

                    if row_idx > 0:
                        matrix_row = t7_obj.matrix[row_idx]
                        year_ready_1cl += float(matrix_row[2].value or 0.0)
                        year_ready_2cl += float(matrix_row[4].value or 0.0)
                        year_ready_3cl += float(matrix_row[6].value or 0.0)
                        if len(matrix_row) > 8:
                            year_ready_4cl += float(matrix_row[8].value or 0.0)

        # Загрузка
        left_1 = BurialSiteDeep.Load("1 класс", year_ready_1cl)
        left_2 = BurialSiteDeep.Load("2 класс", year_ready_2cl)
        left_3 = BurialSiteSurface.Load("3 класс", year_ready_3cl)
        left_4 = BurialSiteSurface.Load("4 класс", year_ready_4cl)

        # Объем который загружен в хранилище
        filled_1cl += (year_ready_1cl - (left_1 or 0.0))
        filled_2cl += (year_ready_2cl - (left_2 or 0.0))
        filled_3cl += (year_ready_3cl - (left_3 or 0.0))
        filled_4cl += (year_ready_4cl - (left_4 or 0.0))

        # Записываем точки для графиков на текущем шаге года
        accumulation_data["years"].append(yr)
        accumulation_data["1 класс (ПГЗР)"].append(filled_1cl)
        accumulation_data["2 класс (ПГЗР)"].append(filled_2cl)
        accumulation_data["3 класс (ППЗР)"].append(filled_3cl)
        accumulation_data["4 класс (ППЗР)"].append(filled_4cl)
    return accumulation_data


class InterfaceRenderer:
    @staticmethod
    def RenderTableUI(active_table_obj, table_field_name, active_id, start_yr, end_yr, technology, selected_year,
                      plant):
        if active_table_obj is None:
            st.info("Данные отсутствуют или папка сценария не задана.")
            return

        # Получаем DataFrame из матрицы объектов Cell
        df_display = active_table_obj.ToDataFrame()
        if df_display.empty:
            st.info("Выбранная таблица пуста.")
            return
        if table_field_name in ["t1", "t2", "t3", "t4"]:
            storage_key = f"edited_df_{table_field_name}_{active_id}"
            if storage_key not in st.session_state:
                st.session_state[storage_key] = df_display
            # редактор ячеек в сессии
            edited_df = st.data_editor(
                st.session_state[storage_key],
                use_container_width=True,
                hide_index=True,
                height=380,
                key=f"editor_action_{table_field_name}_{active_id}"
            )

            if not edited_df.equals(st.session_state[storage_key]):
                with st.spinner("Пересчет моделей по новым значениям ячеек..."):
                    active_table_obj.UpdateFromDataFrame(old_df=st.session_state[storage_key], new_df=edited_df)
                    st.session_state[storage_key] = edited_df
                    plant.CalculateMatrices(technology=technology, start_year=start_yr, end_year=end_yr)
                    st.toast("Матрицы успешно обновлены в памяти!")
                    st.rerun()
        else:
            st.dataframe(
                df_display, use_container_width=True, hide_index=True, height=380,
                key=f"df_{table_field_name}_{start_yr}_{end_yr}_{technology}_{selected_year}_{active_id}"
            )

    @staticmethod
    def RenderConstructorUI(start_yr, end_yr, current_config, active_id, technology):
        import streamlit as st
        import os
        from app import load_base_excel_files
        st.markdown("###Конструктор отчета Word (.docx)")
        st.caption("Выберите элементы программы, настройте их масштаб и порядок для записи в итоговый документ.")
        constructor = ReportConstructor()
        available_items = constructor.GetAvailableElements(st.session_state.plants_data, start_yr, end_yr)
        with st.container(border=True):
            st.markdown("**Шаг 1. Добавить элемент в документ**")
            item_options = {item["label"] + (f" ({item['group']})" if item['plant_name'] else ""): item for item in
                            available_items}
            selected_item_label = st.selectbox("Выберите таблицу или график для вставки:",
                                               options=list(item_options.keys()))
            element_size = st.slider("Ширина элемента в документе (дюймы):", min_value=2.0, max_value=8.0, value=5.5,
                                     step=0.5)
            if st.button("Добавить выбранный элемент в очередь отчета", use_container_width=True):
                chosen_meta = item_options[selected_item_label]
                st.session_state.constructor_queue.append({
                    "key": chosen_meta["key"], "type": chosen_meta["type"], "label": chosen_meta["label"],
                    "field_name": chosen_meta["field_name"], "plant_name": chosen_meta["plant_name"],
                    "tech": chosen_meta.get("tech"), "year": chosen_meta.get("year"), "size": element_size
                })
                st.toast("Элемент добавлен в очередь отчета!")
                st.rerun()

        st.markdown("**Шаг 2. Текущая структура документа (Очередь сборки)**")
        if not st.session_state.constructor_queue:
            st.info("Очередь пуста. Добавьте элементы с помощью формы выше.")
        else:
            for idx, q_item in enumerate(st.session_state.constructor_queue):
                q_col_text, q_col_size, q_col_del = st.columns([6.0, 2.0, 1.0])
                with q_col_text:
                    prefix = "📊" if q_item["type"] == "table" else "📈"
                    plant_prefix = f"[{q_item['plant_name']}] " if q_item["plant_name"] else "[Глобальный] "
                    st.write(f"{idx + 1}. {prefix} {plant_prefix}{q_item['label']}")
                with q_col_size:
                    st.caption(f"Ширина: {q_item['size']} дюйм.")
                with q_col_del:
                    if st.button("❌", key=f"del_q_{idx}"):
                        st.session_state.constructor_queue.pop(idx)
                        st.rerun()

            st.divider()

            btn_build_col, btn_cancel_col = st.columns(2)
            with btn_build_col:
                if st.button("Сгенерировать и сохранить отчет Word", use_container_width=True, type="primary"):
                    with st.spinner("Идет генерация документа конструктором..."):
                        from office import ReportDocument
                        doc_report = ReportDocument(title="Аналитический комплекс: Пользовательский отчет")
                        saved_file = doc_report.BuildDocumentFromQueue(
                            constructor_queue=st.session_state.constructor_queue,
                            plants_session_data=st.session_state.plants_data,
                            start_yr=start_yr, end_yr=end_yr, current_config=current_config
                        )
                        st.session_state.report_status = f"Пользовательский отчет '{os.path.basename(saved_file)}' успешно собран!"
                        st.session_state.constructor_queue = []
                        st.session_state.report_mode = False
                        st.rerun()

            with btn_cancel_col:
                if st.button("↩Закрыть конструктор (Вернуться к таблицам)", use_container_width=True):
                    st.session_state.report_mode = False
                    st.rerun()


class InterfaceRenderer:
    @staticmethod
    def RenderTableUI(active_table_obj, table_field_name, active_id, start_yr, end_yr, technology, selected_year,
                      plant):
        import streamlit as st

        if active_table_obj is None:
            st.info("Данные отсутствуют или папка сценария не задана.")
            return

        # Получаем чистый DataFrame из матрицы объектов Cell
        df_display = active_table_obj.ToDataFrame()

        if df_display.empty:
            st.info("Выбранная таблица пуста.")
            return

        # Разделяем просмотр и корректировку данных
        if table_field_name in ["t1", "t2", "t3", "t4"]:
            storage_key = f"edited_df_{table_field_name}_{active_id}"

            if storage_key not in st.session_state:
                st.session_state[storage_key] = df_display

            # Интерактивный редактор ячеек в памяти сессии
            edited_df = st.data_editor(
                st.session_state[storage_key],
                use_container_width=True,
                hide_index=True,
                height=380,
                key=f"editor_action_{table_field_name}_{active_id}"
            )

            if not edited_df.equals(st.session_state[storage_key]):
                with st.spinner("Пересчет моделей по новым значениям ячеек..."):
                    active_table_obj.UpdateFromDataFrame(old_df=st.session_state[storage_key], new_df=edited_df)
                    st.session_state[storage_key] = edited_df
                    plant.CalculateMatrices(technology=technology, start_year=start_yr, end_year=end_yr)
                    st.toast("Матрицы успешно обновлены в памяти!")
                    st.rerun()
        else:
            st.dataframe(
                df_display, use_container_width=True, hide_index=True, height=380,
                key=f"df_{table_field_name}_{start_yr}_{end_yr}_{technology}_{selected_year}_{active_id}"
            )

    @staticmethod
    def RenderConstructorUI(start_yr, end_yr, current_config, active_id, technology):
        import streamlit as st
        import os
        from logic import ReportConstructor

        st.markdown("### Конструктор отчета Word (.docx)")
        st.caption("Выберите элементы программы, настройте их масштаб и порядок для записи в итоговый документ.")

        constructor = ReportConstructor()
        available_items = constructor.GetAvailableElements(st.session_state.plants_data, start_yr, end_yr)

        with st.container(border=True):
            st.markdown("**Шаг 1. Добавить элемент в документ**")
            item_options = {item["label"] + (f" ({item['group']})" if item['plant_name'] else ""): item for item in
                            available_items}
            selected_item_label = st.selectbox("Выберите таблицу или график для вставки:",
                                               options=list(item_options.keys()))
            element_size = st.slider("Ширина элемента в документе (дюймы):", min_value=2.0, max_value=8.0, value=5.5,
                                     step=0.5)

            if st.button("➕ Добавить выбранный элемент в очередь отчета", use_container_width=True):
                chosen_meta = item_options[selected_item_label]
                st.session_state.constructor_queue.append({
                    "key": chosen_meta["key"], "type": chosen_meta["type"], "label": chosen_meta["label"],
                    "field_name": chosen_meta["field_name"], "plant_name": chosen_meta["plant_name"],
                    "tech": chosen_meta.get("tech"), "year": chosen_meta.get("year"), "size": element_size
                })
                st.toast("Элемент добавлен в очередь отчета!")
                st.rerun()

        st.markdown("**Шаг 2. Текущая структура документа (Очередь сборки)**")
        if not st.session_state.constructor_queue:
            st.info("Очередь пуста. Добавьте элементы с помощью формы выше.")
        else:
            for idx, q_item in enumerate(st.session_state.constructor_queue):
                q_col_text, q_col_size, q_col_del = st.columns([6.0, 2.0, 1.0])
                with q_col_text:
                    prefix = "📊" if q_item["type"] == "table" else "📈"
                    plant_prefix = f"[{q_item['plant_name']}] " if q_item["plant_name"] else "[Глобальный] "
                    st.write(f"{idx + 1}. {prefix} {plant_prefix}{q_item['label']}")
                with q_col_size:
                    st.caption(f"Ширина: {q_item['size']} дюйм.")
                with q_col_del:
                    if st.button("❌", key=f"del_q_{idx}"):
                        st.session_state.constructor_queue.pop(idx)
                        st.rerun()

            st.divider()

            btn_build_col, btn_cancel_col = st.columns(2)
            with btn_build_col:
                if st.button("Сгенерировать и сохранить отчет Word", use_container_width=True, type="primary"):
                    with st.spinner("Идет генерация документа конструктором..."):
                        from office import ReportDocument
                        doc_report = ReportDocument(title="Аналитический комплекс: Пользовательский отчет")
                        saved_file = doc_report.BuildDocumentFromQueue(
                            constructor_queue=st.session_state.constructor_queue,
                            plants_session_data=st.session_state.plants_data,
                            start_yr=start_yr, end_yr=end_yr, current_config=current_config
                        )
                        st.session_state.report_status = f"Пользовательский отчет '{os.path.basename(saved_file)}' успешно собран!"
                        st.session_state.constructor_queue = []
                        st.session_state.report_mode = False
                        st.rerun()

            with btn_cancel_col:
                if st.button("↩Закрыть конструктор (Вернуться к таблицам)", use_container_width=True):
                    st.session_state.report_mode = False
                    st.rerun()

    @staticmethod
    def RenderConstantsUI(active_id, technology, start_yr, end_yr, plant):
        import streamlit as st

        with st.container(border=True):
            st.markdown("** Корректировка констант**")
            st.caption("Изменение порогов тепловыделения и периодов полураспада РАО по классам.")
            classes = ["1 класс", "2 класс", "3 класс"]

            # Cетка полей для критериев тепловыделения (FuelClassHeatReady)
            st.markdown("*Критерии тепловыделения (кВт/м³):*")
            c_cols1 = st.columns(2)
            for idx, cl in enumerate(classes):
                with c_cols1[idx % 2]:
                    new_heat = st.number_input(
                        f"{cl}",
                        value=float(st.session_state.FuelClassHeatReady[cl]),
                        step=0.1,
                        format="%.1f",
                        key=f"const_heat_{cl}_{active_id}"
                    )
                    # Если пользователь поменял число, мгновенно фиксируем в словаре
                    if new_heat != st.session_state.FuelClassHeatReady[cl]:
                        st.session_state.FuelClassHeatReady[cl] = new_heat
            st.markdown("*Периоды полураспада (годы):*")
            c_cols2 = st.columns(2)
            for idx, cl in enumerate(classes):
                with c_cols2[idx % 2]:
                    new_hl = st.number_input(
                        f"{cl}",
                        value=float(st.session_state.HalfLife[cl]),
                        step=0.5,
                        format="%.1f",
                        key=f"const_hl_{cl}_{active_id}"
                    )
                    # Фиксируем изменения полураспада в словаре сессии
                    if new_hl != st.session_state.HalfLife[cl]:
                        st.session_state.HalfLife[cl] = new_hl

            # Кнопка применения изменений и сквозного пересчета матриц
            if st.button("Применить константы", use_container_width=True, type="primary",
                         key=f"apply_const_btn_{active_id}"):
                with st.spinner("Пересчет моделей отрасли по новым константам..."):
                    if plant is not None:
                        # Принудительно взводим триггер пересчета и перезапускаем CalculateMatrices
                        st.session_state[f"need_calc_{active_id}"] = True
                        plant.CalculateMatrices(technology=technology, start_year=start_yr, end_year=end_yr)

                    st.toast("Константы успешно обновлены! Модели Т7 пересчитаны.")
                    st.rerun()
def GetCapacityAnalysisData(plants_session_dict, start_yr, end_yr=2050):
    analysis_data = {"years": [], "t1_total": [], "t2_total_sum": []}
    # ограничиваем 2050 годом, пока Завод 2 не продлен
    final_end_year = min(end_yr, 2050)

    # Берем первый попавшийся завод, чтобы дотянуться до общей таблицы Т1
    first_plant = None
    for p_id in plants_session_dict.keys():
        import streamlit as st
        if f"plant_object_{p_id}" in st.session_state:
            first_plant = st.session_state[f"plant_object_{p_id}"]
            break
    for yr in range(start_yr, final_end_year + 1):
        analysis_data["years"].append(yr)
        t1_sum = 0.0
        if first_plant is not None and first_plant.t1 is not None:
            # Ищем строку года в Т1 (год в колонке 0)
            for r in range(1, len(first_plant.t1.matrix)):
                if first_plant.t1.matrix[r][0].value == yr:
                    # Суммируем значения по всем реакторам
                    for c in range(1, len(first_plant.t1.matrix[r])):
                        val = first_plant.t1.matrix[r][c].value
                        t1_sum += float(val) if val is not None else 0.0
                    break
        analysis_data["t1_total"].append(t1_sum)
        t2_plants_sum = 0.0
        for p_id, p_info in plants_session_dict.items():
            import streamlit as st
            plant_key = f"plant_object_{p_id}"
            if plant_key in st.session_state and st.session_state[plant_key] is not None:
                p_obj = st.session_state[plant_key]
                if p_obj.t2 is not None:
                    for r in range(1, len(p_obj.t2.matrix)):
                        if p_obj.t2.matrix[r][0].value == yr:
                            # Суммируем загрузку всех реакторов на этом заводе (колонки 1-5)
                            for c in range(2, len(first_plant.t1.matrix[r])):
                                val = first_plant.t1.matrix[r][c].value
                                t1_sum += float(val) if not math.isnan(val) else 0.0
                            break
        analysis_data["t2_total_sum"].append(t2_plants_sum)
    return analysis_data
def GetBurialVsCapacityData(plants_session_dict, start_yr, end_yr=2050):
    analysis_data = {"years": [], "total_t2": [], "total_burial_ready": []}
    final_end_year = min(end_yr, 2050)
    for yr in range(start_yr, final_end_year + 1):
        analysis_data["years"].append(yr)
        t2_sum = 0.0
        burial_sum = 0.0

        for p_id, p_info in plants_session_dict.items():
            import streamlit as st
            plant_key = f"plant_object_{p_id}"
            if plant_key in st.session_state and st.session_state[plant_key] is not None:
                p_obj = st.session_state[plant_key]
                tech = p_info["technology"]

                # Извлекаем Т2 загрузку
                if p_obj.t2 is not None:
                    for r in range(1, len(p_obj.t2.matrix)):
                        if p_obj.t2.matrix[r][0].value == yr:
                            for c in range(1, len(p_obj.t2.matrix[r])):
                                val = p_obj.t2.matrix[r][c].value
                                t2_sum += float(val) if val is not None else 0.0
                            break
                # Извлекаем Т7 готовые кубометры отходов
                t7_dict = getattr(p_obj, "t7", {})
                t7_obj = t7_dict.get(tech) if isinstance(t7_dict, dict) else None
                if t7_obj is not None and hasattr(t7_obj, 'matrix'):
                    for r in range(1, len(t7_obj.matrix)):
                        if t7_obj.matrix[r][0].value == yr:
                            matrix_row = t7_obj.matrix[r]
                            # Четные колонки в Т7: 2 (1кл), 4 (2кл), 6 (3кл) и 8 (4кл) — это объемы "Готово"
                            columns_to_check = [2, 4, 6]
                            if len(matrix_row) > 8:
                                columns_to_check.append(8)

                            for col_idx in columns_to_check:
                                val = matrix_row[col_idx].value
                                burial_sum += float(val) if val is not None else 0.0
                            break
        analysis_data["total_t2"].append(t2_sum)
        analysis_data["total_burial_ready"].append(burial_sum)
    return analysis_data