import pandas as pd
import docx
from logic import *
import os
import io
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from docx import Document
from docx.shared import Inches
from logic import Graph, NuclearDataVisualizer

class Word:
    def __init__(self, filename):
        self.FileName = filename
        self.Extension = 'docx'
        self.FullName = f'{self.FileName}.{self.Extension}'

        if os.path.exists(self.FullName):
            self.doc = docx.Document(self.FullName)
        else:
            self.doc = docx.Document()
            self.doc.save(self.FullName)

    def Save(self):
        self.doc.save(self.FullName)

    def ReadFileText(self):
        text = [par.text for par in self.doc.paragraphs]
        return '\n'.join(text)

    def ReadFileTable(self, table_number):
        if table_number >= len(self.doc.tables):
            return "Таблица не найдена"
        table = self.doc.tables[table_number]
        return [[cell.text for cell in row.cells] for row in table.rows]

    def AddHeader(self, header="Заголовок", level=1):
        self.doc.add_heading(header, level=level)
        self.Save()

    def AddText(self, text='Текст.'):
        self.doc.add_paragraph(text)
        self.Save()

    def AddTable(self, table_obj):
        rows = len(table_obj.matrix)
        cols = len(table_obj.matrix[0])
        word_table = self.doc.add_table(rows=rows, cols=cols)


def ToExcel(filename, table, sheetname=False):
    df = pd.DataFrame(table)
    if not sheetname:
        df.to_excel(f'excel/{filename}.xlsx', index=False)
    else:
        with pd.ExcelWriter(f'excel/{filename}.xlsx') as writer:
            df.to_excel(writer, sheet_name=sheetname)
    return None


import os
import io
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from docx import Document
from docx.shared import Inches
from logic import Graph, NuclearDataVisualizer


def CreateWordReport(plant=None, sensitivity_data=None, tech_num=1):
    global tables, graphs_t1, graphs_t2, reactors

    reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']
    target_years = [2030, 2035, 2040, 2045, 2050]

    if plant is not None:
        plant.calculate_matrices(technology=1, start_year=2030, end_year=2050)
        plant.calculate_matrices(technology=2, start_year=2030, end_year=2050)

        tables = {
            't1': plant.t1,
            't2': plant.t2,
            't3': plant.t3,
            't4': plant.t4,
            't5-1': plant.t5_series[1],
            't5-2': plant.t5_series[2],
            't6-1': plant.t6_series[1],
            't6-2': plant.t6_series[2],
            't7': plant.t7[tech_num]
        }

        graphs_t1 = {}
        graphs_t2 = {}
        for i, fuel in enumerate(reactors):
            graphs_t1[fuel] = Graph(table=plant.t1, column=i + 1)
            graphs_t2[fuel] = Graph(table=plant.t2, column=i + 1)

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

    def add_plot_to_doc(vis_obj, title, is_sensitivity=False):
        if not is_sensitivity and not vis_obj.datasets:
            return

        buf = io.BytesIO()
        plt.figure(figsize=(10, 5))

        if is_sensitivity and sensitivity_data:
            x_percent, y_volumes = sensitivity_data
            plt.plot(x_percent, y_volumes, marker='s', color='#d9534f', linewidth=2, linestyle='-')
            if 100 in x_percent:
                nominal_idx = x_percent.index(100)
                plt.plot(100, y_volumes[nominal_idx], marker='o', color='black', markersize=8, label='Номинал (100%)')
            plt.xlabel("Изменение начального тепловыделения РАО 1 класса, %", fontsize=10)
            plt.ylabel("Готово к захоронению к 2050 году, м³", fontsize=10)
            plt.gca().xaxis.set_major_locator(MultipleLocator(10))
        else:
            for label, (x, y) in vis_obj.datasets.items():
                plt.plot(x, y, marker='o', label=label, linewidth=2, markersize=4)
            plt.xlabel("Годы")
            plt.ylabel("Значение показателей")
            plt.gca().xaxis.set_major_locator(MultipleLocator(5))

        plt.title(title, fontsize=12)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
        plt.tight_layout()

        plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        doc.add_picture(buf, width=Inches(5.5))
        plt.close()
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

    for t_key in ['t3', 't4']:
        if t_key not in tables:
            continue
        table_obj = tables[t_key]

        title_text = "Нормативы образования РАО (Т3)" if t_key == 't3' else "Удельное тепловыделение РАО (Т4)"
        doc.add_heading(title_text, level=1)
        doc.add_paragraph(getattr(table_obj, 'description', ''))

        raw_matrix = [[cell.value for cell in row] for row in table_obj.matrix]
        if len(raw_matrix) == 0:
            continue

        rows, cols = len(raw_matrix), len(raw_matrix[0])
        word_tab = doc.add_table(rows=rows, cols=cols)
        word_tab.style = 'Table Grid'

        for r in range(rows):
            for c in range(cols):
                cell_val = raw_matrix[r][c]

                if cell_val is None or str(cell_val).lower() == "nan" or str(cell_val) == "None":
                    text_to_write = ""
                else:
                    if isinstance(cell_val, (int, float)):
                        text_to_write = str(round(cell_val, 2))
                    else:
                        text_to_write = str(cell_val)

                cell = word_tab.cell(r, c)
                cell.text = text_to_write

                if r == 0:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True

        doc.add_page_break()

    for t_num in [1, 2]:
        for t_base in ['t5', 't6']:
            t_code = f'{t_base}-{t_num}'

            # Если данных по этой технологии нет в текущем расчете - пропускаем
            if t_code not in tables or not tables[t_code]:
                continue

            doc.add_heading(f'Результаты расчетов {t_base.upper()} (Технология {t_num})', level=1)
            for yr in target_years:
                if yr not in tables[t_code]:
                    continue
                doc.add_heading(f'Год {yr}', level=2)
                curr_t = tables[t_code][yr]

                rows, cols = len(curr_t.matrix), len(curr_t.matrix[0])
                word_tab = doc.add_table(rows=rows, cols=cols)
                word_tab.style = 'Table Grid'
                for r in range(rows):
                    for c in range(cols):
                        word_tab.cell(r, c).text = str(curr_t.matrix[r][c].value)

            vis_s = NuclearDataVisualizer()
            if 't5' in t_base:
                for r in reactors:
                    g_obj = Graph.FromTableSeries(tables[t_code], r, '1 класс')
                    if g_obj.graph:
                        vis_s.add_graph(r, g_obj)
                add_plot_to_doc(vis_s, f"Образование РАО 1 класса ({t_code.upper()})")
            else:
                for cl in ['1 класс', '2 класс', '3 класс']:
                    data_t6 = {}
                    for yr, tab in tables[t_code].items():
                        try:
                            c_idx = tab.FindValue(cl)['column']
                            data_t6[yr] = tab.matrix[c_idx].value
                        except:
                            continue
                    if data_t6:
                        g_t6 = Graph()
                        g_t6.graph = data_t6
                        vis_s.add_graph(cl, g_t6)
                add_plot_to_doc(vis_s, f"Среднее тепловыделение ({t_code.upper()})")

            doc.add_page_break()

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

    # Анализ чувствительности
    if sensitivity_data:
        doc.add_page_break()
        doc.add_heading('Анализ чувствительности объемов РАО', level=1)
        doc.add_paragraph(
            "В данном разделе представлены результаты исследования чувствительности суммарного объема "
            "РАО 1 класса, готового к захоронению к 2050 году, при варьировании исходных показателей "
            "удельного тепловыделения в диапазоне от 50% до 150% от номинальных значений."
        )
        add_plot_to_doc(None, f"Зависимость объемов РАО 1 класса к 2050 году от тепловыделения (Технология {tech_num})", is_sensitivity=True)

    doc.save(filename)
    print(f"Готово! Отчет сохранен: {filename}")