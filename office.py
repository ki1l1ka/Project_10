import pandas as pd
import docx

from logic import *
import os

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

report = Word("Отчёт")
report.AddHeader("Заг 1")
report.AddText("Вот это вот все текст")
print(report.ReadFileText())
report.AddTable(Table(TableExcel('t1'), description="Таблица T1 Исходные данные про образованию ОЯТ по годам, тонны тяжелого металла"))
print(report.ReadFileTable(0))