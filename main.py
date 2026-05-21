from logic import *
import os
import io
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from docx import Document
from docx.shared import Inches
tables = {'t1': Table(TableExcel('t1'), description="Таблица T1 Исходные данные про образованию ОЯТ по годам, тонны тяжелого металла"),
          't2': Table(TableExcel('t2'), description="Таблица T2 Исходные данные по загрузке завода 1 ОЯТ различных типов (тонны тяжелого металла) по годам"),
          't3': Table(TableExcel('t3'), description="Таблица T3 Исходные данные. Образование РАО различных классов (кубометры РАО на тонну тяжелого металла) при переработке ОЯТ различных типов по различным технологиям"),
          't4': Table(TableExcel('t4'), description="Таблица T4 Исходные данные. Тепловыделение РАО сразу после переработки (кВт на кубический метр)")}
headers = {'t5': ['Завод',	'1 класс', '2 класс', '3 класс', '4 класс'],
           't6': ['Завод', '1 класс', '2 класс', '3 класс','4 класс'],
           't7': ["Год", '1 класс', "Готово к звхоронению", '2 класс', "Готово к звхоронению", '3 класс', "Готово к звхоронению"]}
lines = {'t5': ['Завод', 'ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК'],
         't6': ['Завод', '...'],
         't7': ["Год", 2030]}

tables['t5-1'] = dict()
for i in range(2030, 2051):
    tables['t5-1'][i] = CreatedTable(TableCreate(headers['t5'], lines['t5']))
    tables['t5-1'][i].T5Create(t2=tables['t2'], t3=tables['t3'], t5=tables['t5-1'][i], year=i, technology=1)
tables['t6-1'] = dict()
for i in range(2030, 2051):
    tables['t6-1'][i] = CreatedTable(TableCreate(headers['t6'], lines['t6']))
    tables['t6-1'][i].T6Create(table6=tables['t6-1'][i], tables5=tables['t5-1'], table4=tables['t4'], year=i, technology=1)

tables['t5-2'] = dict()
for i in range(2030, 2051):
    tables['t5-2'][i] = CreatedTable(TableCreate(headers['t5'], lines['t5']))
    tables['t5-2'][i].T5Create(t2=tables['t2'], t3=tables['t3'], t5=tables['t5-2'][i], year=i, technology=2)

tables['t6-2'] = dict()
for i in range(2030, 2051):
    tables['t6-2'][i] = CreatedTable(TableCreate(headers['t6'], lines['t6']))
    tables['t6-2'][i].T6Create(table6=tables['t6-2'][i], tables5=tables['t5-2'], table4=tables['t4'], year=i, technology=2)

tables['t7'] = CreatedTable(TableCreate(headers['t7']), description="Количество РАО по годам накопительным итогом, в том числе готово к захоронению  накопленных за год (кубометры)")
tables['t7'].T7Create(tables['t4'], tables['t5-1'], tables['t7'])
graphs_t7_by_fuel = {}
for fuel in ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']:
    graphs_t7_by_fuel[fuel] = Graph.FromTableSeries(tables['t5-1'], fuel, '1 класс')
reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']

graphs_t1 = {}
graphs_t2 = {}

for i, fuel in enumerate(reactors):
    # В t1 и t2 топливо идет в колонках (индексы 1-5), годы в строках
    graphs_t1[fuel] = Graph(table=tables['t1'], column=i + 1)
    graphs_t2[fuel] = Graph(table=tables['t2'], column=i + 1)

# --- Графики для T5 (Серия таблиц) ---
# Создадим графики для 1-го класса (можно сделать цикл по всем классам)
graphs_t5_1 = {}
for fuel in reactors:
    graphs_t5_1[fuel] = Graph.FromTableSeries(tables['t5-1'], fuel, '1 класс')

# --- Графики для T6 (Среднее тепловыделение) ---
# В T6 у вас структура ['Завод', '1 класс', ...], берем данные из второй строки (индекс 1)
graphs_t6_1 = {}
for cl in ['1 класс', '2 класс', '3 класс']:
    # Здесь fuel_type не нужен, так как в T6 всего одна строка с данными
    # Но метод FromTableSeries можно адаптировать или использовать цикл:
    data = {}
    for year, table in tables['t6-1'].items():
        col_idx = table.FindValue(cl)['column']
        data[year] = table.matrix[1][col_idx].value

    g_obj = Graph()
    g_obj.graph = data
    graphs_t6_1[cl] = g_obj

# Графики для T7
graphs_t7 = {
    '1 класс (Всего)': Graph(table=tables['t7'], column=1),
    '1 класс (Готово)': Graph(table=tables['t7'], column=2),
    '2 класс (Всего)': Graph(table=tables['t7'], column=3),
    '2 класс (Готово)': Graph(table=tables['t7'], column=4)
}
# Визуализаторы для базовых таблиц
vis_t1 = TableVisualizer(tables['t1'])
vis_t2 = TableVisualizer(tables['t2'])
vis_t4 = TableVisualizer(tables['t4'])

# Визуализаторы для рассчитанных серий
vis_t5_2030 = TableVisualizer(tables['t5-1'][2030])
vis_t6_2030 = TableVisualizer(tables['t6-1'][2030])

# Визуализатор для итоговой таблицы t7
vis_t7 = TableVisualizer(tables['t7'])

def CreateWordReport(sensitivity_data=None, tech_num=1):
    folder = "отчёты"
    if not os.path.exists(folder):
        os.makedirs(folder)

    n = 1
    while os.path.exists(f"{folder}/отчёт {n}.docx"):
        n += 1
    filename = f"{folder}/отчёт {n}.docx"

    print(f"Начинаю сборку отчета: {filename}")
    doc = Document()
    doc.add_heading('Аналитический отчет по обращению с ОЯТ и РАО', 0)

    target_years = [2030, 2035, 2040, 2045, 2050]
    reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']

    # функция для графиков
    def add_plot_to_doc(vis_obj, title, is_sensitivity=False):
        if not is_sensitivity and not vis_obj.datasets:
            return

        buf = io.BytesIO()
        plt.figure(figsize=(10, 5))

        if is_sensitivity and sensitivity_data:
            # Отрисовка графика чувствительности
            x_percent, y_volumes = sensitivity_data
            plt.plot(x_percent, y_volumes, marker='s', color='#d9534f', linewidth=2, linestyle='-')
            if 100 in x_percent:
                nominal_idx = x_percent.index(100)
                plt.plot(100, y_volumes[nominal_idx], marker='o', color='black', markersize=8, label='Номинал (100%)')
            plt.xlabel("Изменение начального тепловыделения РАО 1 класса, %", fontsize=10)
            plt.ylabel("Готово к захоронению к 2050 году, м³", fontsize=10)
            plt.gca().xaxis.set_major_locator(MultipleLocator(10))  # Шаг 10%
        else:
            # Стандартная отрисовка графиков реакторов Т7
            for label, (x, y) in vis_obj.datasets.items():
                plt.plot(x, y, marker='o', label=label, linewidth=2, markersize=4)
            plt.xlabel("Годы")
            plt.ylabel("Значение показателей")
            plt.gca().xaxis.set_major_locator(MultipleLocator(5))  # Шаг 5 лет

        plt.title(title, fontsize=12)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
        plt.tight_layout()

        plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        doc.add_picture(buf, width=Inches(5.5))
        plt.close()

    # Т1, Т2
    for t_key in ['t1', 't2']:
        if t_key not in tables: continue
        table_obj = tables[t_key]
        doc.add_heading(f'Данные из таблицы {t_key.upper()}', level=1)
        doc.add_paragraph(getattr(table_obj, 'description', ''))

        rows, cols = len(table_obj.matrix), len(table_obj.matrix[0])
        word_tab = doc.add_table(rows=rows, cols=cols)
        word_tab.style = 'Table Grid'
        for r in range(rows):
            for c in range(cols):
                word_tab.cell(r, c).text = str(table_obj.matrix[r][c].value)

        vis = NuclearDataVisualizer()
        source = graphs_t1 if t_key == 't1' else graphs_t2
        for r in reactors:
            if r in source: vis.add_graph(r, source[r])
        add_plot_to_doc(vis, f"Динамика мощностей {t_key.upper()}")
        doc.add_page_break()

    # T5, T6
    for t_code in [f't5-{tech_num}', f't6-{tech_num}']:
        if t_code not in tables: continue
        doc.add_heading(f'Результаты расчетов {t_code.upper()}', level=1)

        for yr in target_years:
            if yr not in tables[t_code]: continue
            doc.add_heading(f'Год {yr}', level=2)
            curr_t = tables[t_code][yr]

            rows, cols = len(curr_t.matrix), len(curr_t.matrix[0])
            word_tab = doc.add_table(rows=rows, cols=cols)
            word_tab.style = 'Table Grid'
            for r in range(rows):
                for c in range(cols):
                    word_tab.cell(r, c).text = str(curr_t.matrix[r][c].value)

        vis_s = NuclearDataVisualizer()
        if 't5' in t_code:
            for r in reactors:
                g_obj = Graph.FromTableSeries(tables[t_code], r, '1 класс')
                if g_obj.graph: vis_s.add_graph(r, g_obj)
            add_plot_to_doc(vis_s, f"Образование РАО 1 класса ({t_code})")
        else:
            for cl in ['1 класс', '2 класс', '3 класс']:
                data_t6 = {}
                for yr, tab in tables[t_code].items():
                    try:
                        c_idx = tab.FindValue(cl)['column']
                        data_t6[yr] = tab.matrix[1][c_idx].value
                    except:
                        continue
                if data_t6:
                    g_t6 = Graph();
                    g_t6.graph = data_t6
                    vis_s.add_graph(cl, g_t6)
            add_plot_to_doc(vis_s, f"Среднее тепловыделение ({t_code})")

    # T7
    if 't7' in tables:
        doc.add_page_break()
        doc.add_heading('Итоговая таблица T7 и анализ готовности РАО', level=1)
        t7 = tables['t7']

        rows_t7, cols_t7 = len(t7.matrix), len(t7.matrix[0])
        word_t7 = doc.add_table(rows=rows_t7, cols=cols_t7)
        word_t7.style = 'Table Grid'
        for r in range(rows_t7):
            for c in range(cols_t7):
                val = t7.matrix[r][c].value
                word_t7.cell(r, c).text = str(val) if val is not None else ""

        vis_t7_compare = NuclearDataVisualizer()
        classes = [("1 класс", 1, 2), ("2 класс", 3, 4), ("3 класс", 5, 6)]
        for name, col_all, col_ready in classes:
            vis_t7_compare.add_graph(f"{name} (Всего)", Graph(table=t7, column=col_all))
            vis_t7_compare.add_graph(f"{name} (Готово)", Graph(table=t7, column=col_ready))
        add_plot_to_doc(vis_t7_compare, "Сравнение накопленных и готовых к захоронению объемов")

    # Чувствительность
    if sensitivity_data:
        doc.add_page_break()
        doc.add_heading('Анализ чувствительности объемов РАО', level=1)
        doc.add_paragraph(
            "В данном разделе представлены результаты исследования чувствительности суммарного объема "
            "РАО 1 класса, готового к захоронению к 2050 году, при варьировании исходных показателей "
            "удельного тепловыделения в диапазоне от 50% до 150% от номинальных значений."
        )
        # Добавляем график чувствительности в документ
        add_plot_to_doc(None, f"Зависимость объемов РАО 1 класса к 2050 году от тепловыделения (Технология {tech_num})",
                        is_sensitivity=True)

    doc.save(filename)
    print(f"Готово! Отчет сохранен: {filename}")


CreateWordReport()