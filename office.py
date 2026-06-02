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

import os
import io
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.shared import Inches


class ReportDocument:
    def __init__(self, title="Аналитический отчет по обращению с ОЯТ и РАО"):
        self.doc = Document()
        self.doc.add_heading(title, 0)

        # Папка для сохранения по умолчанию
        self.folder = "отчёты"
        if not os.path.exists(self.folder):
            os.makedirs(self.folder)

    def AddHeading(self, text, level=1):
        """Добавление текстового заголовка перед блоком"""
        self.doc.add_heading(text, level=level)

    def AddParagraph(self, text):
        """Добавление описания или пояснительного текста"""
        if text:
            self.doc.add_paragraph(text)

    def AddPageBreak(self):
        """Принудительный перенос страницы"""
        self.doc.add_page_break()

    def AddTable(self, table_obj, plant_name=None):
        """
        Нативное добавление очищенной матрицы ячеек любого типа таблицы (Table/CreatedTable).
        Учитывает пустые nan-ячейки и округляет числа.
        """
        if table_obj is None or not hasattr(table_obj, 'matrix'):
            return

        # Парсим матрицу значений ячеек объектов Cell
        raw_matrix = [[cell.value for cell in row] for row in table_obj.matrix]
        if len(raw_matrix) == 0:
            return

        rows = len(raw_matrix)
        cols = max(len(row) for row in raw_matrix)

        # Формируем и выравниваем заголовки
        headers = []
        for cell_val in raw_matrix[0]:
            val_str = str(cell_val).strip() if pd.notna(cell_val) else ""
            if val_str.lower() in ["nan", "none"]:
                val_str = ""
            headers.append(val_str)

        if len(headers) < cols:
            headers.extend([""] * (cols - len(headers)))
        else:
            headers = headers[:cols]

        # Создаем таблицу в документе Word
        word_tab = self.doc.add_table(rows=1, cols=cols)
        word_tab.style = 'Table Grid'

        # Записываем шапку
        hdr_cells = word_tab.rows[0].cells
        for c in range(cols):
            hdr_cells[c].text = headers[c]
            for p in hdr_cells[c].paragraphs:
                for r in p.runs:
                    r.bold = True

        # Записываем тело данных
        for r in range(1, rows):
            row_cells = word_tab.add_row().cells
            for c in range(min(len(raw_matrix[r]), cols)):
                val = raw_matrix[r][c]
                if pd.isna(val) or str(val).lower() in ["nan", "none"]:
                    text_str = ""
                else:
                    text_str = str(round(val, 2)) if isinstance(val, (int, float)) else str(val)
                row_cells[c].text = text_str

    def AddPlot(self, figure, title="", width_inches=5.5):
        """
        Добавление переданного matplotlib холста (figure) в документ
        с возможностью масштабирования по ширине (width_inches)
        """
        if figure is None:
            return

        buf = io.BytesIO()
        # Сохраняем переданную фигуру с обрезкой полей
        figure.savefig(buf, format='png', bbox_inches='tight', dpi=150)

        if title:
            self.doc.add_paragraph(title)

        self.doc.add_picture(buf, width=Inches(width_inches))
        plt.close(figure)

    def Save(self):
        """Фиксация документа на диске с автоинкрементом имени файла"""
        n = 1
        while os.path.exists(f"{self.folder}/отчёт_{n}.docx"):
            n += 1
        filename = f"{self.folder}/отчёт_{n}.docx"

        self.doc.save(filename)
        return filename
