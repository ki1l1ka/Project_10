
from logic import *
import os
import io
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
from docx import Document
from docx.shared import Inches
from logic import *
from office import *

technology_id = 1
start_year = 2030
end_year = 2050

plant1 = ProcessingPlant(name="Завод №1", folder_path="excel/Завод 1")
plant1.calculate_matrices(technology=technology_id, start_year=start_year, end_year=end_year)

tables = {
    't1': plant1.t1, 't2': plant1.t2, 't3': plant1.t3, 't4': plant1.t4,
    f't5-{technology_id}': plant1.t5_series,
    f't6-{technology_id}': plant1.t6_series, 't7': plant1.t7
}

reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']
graphs_t1 = {r: Graph(table=plant1.t1, column=i+1) for i, r in enumerate(reactors)}
graphs_t2 = {r: Graph(table=plant1.t2, column=i+1) for i, r in enumerate(reactors)}
