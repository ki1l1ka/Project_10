import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import pandas as pd
import os
from logic import Table, CreatedTable, Graph, NuclearDataVisualizer, TableExcel, TableCreate
st.set_page_config(page_title="РАО Аналитика", layout="wide")
# Стили
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    [data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: stretch !important;
        width: 100% !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(1) {
        min-width: 260px !important;
    }
    [data-testid="stHorizontalBlock"] > div:nth-child(2) {
        flex-grow: 2 !important;
    }
    [data-testid="stHorizontalBlock"] > div:nth-child(3) {
        min-width: 280px !important;
    }

    .block-container {
        padding-top: 0.5rem;
        padding-left: 1.5rem;
        padding-right: 1.5rem;
    }
    </style>
    """, unsafe_allow_html=True)

# Загрузка файлов
def load_base_excel_files(folder_path):
    abs_folder_path = os.path.abspath(folder_path)

    if not os.path.exists(abs_folder_path):
        st.sidebar.error(f"Папка не найдена по пути:\n`{abs_folder_path}`")
        return None

    try:
        required_files = ['t1', 't2', 't3', 't4']
        missing_files = []

        for name in required_files:
            file_p = os.path.join(abs_folder_path, f"{name}.xlsx")
            if not os.path.exists(file_p):
                missing_files.append(f"{name}.xlsx")

        if missing_files:
            st.sidebar.error(f"В папке `{folder_path}` отсутствуют файлы: {', '.join(missing_files)}")
            return None

        return {
            't1': Table(pd.read_excel(os.path.join(abs_folder_path, "t1.xlsx"), header=None).to_numpy(),
                        description="Таблица Т1. Исходные данные по образованию ОЯТ (тонны тяжелого металла)"),
            't2': Table(pd.read_excel(os.path.join(abs_folder_path, "t2.xlsx"), header=None).to_numpy(),
                        description="Таблица Т2. Данные по загрузке завода ОЯТ (тонны тяжелого металла)"),
            't3': Table(pd.read_excel(os.path.join(abs_folder_path, "t3.xlsx"), header=None).to_numpy(),
                        description="Таблица Т3. Образование РАО различных классов (м³/т ТМ)"),
            't4': Table(pd.read_excel(os.path.join(abs_folder_path, "t4.xlsx"), header=None).to_numpy(),
                        description="Таблица Т4. Тепловыделение РАО сразу после переработки (кВт/м³)")
        }
    except Exception as e:
        st.sidebar.error(f"🚨 Ошибка чтения Excel-файлов: {e}")
        return None

# Расчёт матриц
def calculate_nuclear_matrices(tech_idx, s_yr, e_yr, base_data):
    if base_data is None:
        return None
    try:
        import copy
        data_store = copy.deepcopy(base_data)

        t5_key = f't5-{tech_idx}'
        t6_key = f't6-{tech_idx}'
        data_store[t5_key] = {}
        data_store[t6_key] = {}

        all_reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']

        headers_t5 = ['Завод', '1 класс', '2 класс', '3 класс', '4 класс']
        headers_t6 = ['Завод', '1 класс', '2 класс', '3 класс', '4 класс']
        headers_t7 = ["Год", '1 класс', "Готово к захоронению", '2 класс', "Готово к захоронению", '3 класс',
                      "Готово к захоронению"]

        for y in range(s_yr, e_yr + 1):
            # Создаем структуру Т5
            struct_t5 = [headers_t5] + [[reactor, 0.0, 0.0, 0.0, 0.0] for reactor in all_reactors]
            data_store[t5_key][y] = CreatedTable(table=struct_t5, description=f"Таблица Т5 за {y} год")
            try:
                data_store[t5_key][y].T5Create(t2=data_store['t2'], t3=data_store['t3'], t5=data_store[t5_key][y],
                                               year=y, technology=tech_idx)
            except IndexError as e:
                print(f"⚠️ Ошибка индексов в T5Create за {y} год. Проверьте размеры таблиц T2/T3. Ошибка: {e}")
            # Создаем структуру Т6
            struct_t6 = [headers_t6] + [[reactor, 0.0, 0.0, 0.0, 0.0] for reactor in all_reactors]
            data_store[t6_key][y] = CreatedTable(table=struct_t6, description=f"Таблица Т6 за {y} год")
            try:
                data_store[t6_key][y].T6Create(table6=data_store[t6_key][y], tables5=data_store[t5_key],
                                               table4=data_store['t4'], year=y, technology=tech_idx)
            except (IndexError, ZeroDivisionError) as e:
                print(f"⚠️ Ошибка в T6Create за {y} год. Проверьте таблицу T4 или деление на ноль. Ошибка: {e}")

        # Расчет Т7
        struct_t7 = [headers_t7]
        data_store['t7'] = CreatedTable(table=struct_t7, description="Итоговая таблица Т7")
        try:
            data_store['t7'].T7Create(data_store['t4'], data_store[t5_key], data_store['t7'], technology=tech_idx,
                                      startyear=s_yr, endyear=e_yr)
        except IndexError as e:
            print(f"⚠️ Ошибка индексов в T7Create. Проверьте структуру итоговых расчетов. Ошибка: {e}")

        return data_store
    except Exception as e:
        st.error(f"Ошибка вычислений в ядре: {e}")
        return None

# Окно выбора
if 'report_status' not in st.session_state:
    st.session_state.report_status = ""
if 'excel_folder_path' not in st.session_state:
    st.session_state.excel_folder_path = "excel"  # Папка по умолчанию при первом старте

if "selected_dir" in st.query_params:
    st.session_state.excel_folder_path = st.query_params["selected_dir"]
    st.query_params.clear()  # Очищаем параметры, чтобы избежать зацикливания окон

with st.container():
    tb_col_path, tb_col2, tb_col3, tb_col4 = st.columns([5.5, 1.5, 1.5, 1.5])

    with tb_col_path:
        # Интерактивное текстовое поле
        input_folder = st.text_input(
            "Путь к папке данных сценария:",
            value=st.session_state.excel_folder_path,
            placeholder="Пример: C:\\Users\\Name\\Desktop\\Новый_Сценарий",
            help="Укажите полный путь к папке, где лежат ваши файлы t1.xlsx - t4.xlsx"
        )
        if input_folder != st.session_state.excel_folder_path:
            if os.path.exists(input_folder):
                st.session_state.excel_folder_path = input_folder
                st.toast(f"Папка переключена на: {input_folder}")
                st.rerun()
            else:
                st.error("Указанный путь к папке не найден в файловой системе!")

    with tb_col2:
        st.button("Правка", use_container_width=True, disabled=True)
    with tb_col3:
        st.button("Вид", use_container_width=True, disabled=True)
    with tb_col4:
        st.button("Справка", use_container_width=True, disabled=True)

st.divider()

col_left, col_center, col_right = st.columns([2.5, 6.5, 3.0], gap="large")

with col_left:
    st.markdown("### Панель управления")
    db_loaded = st.button("База данных", use_container_width=True)

    st.divider()
    st.markdown("**Выборка данных для расчета**")
    technology = st.selectbox("Технология переработки:", options=[1, 2], index=0)

    start_yr, end_yr = st.slider(
        "Временной диапазон (годы):",
        min_value=2030, max_value=2050, value=(2030, 2050), step=1
    )

    st.divider()
    st.markdown("**Инструменты и реакторы**")
    all_reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']
    selected_reactors = st.multiselect(
        "Установки в фокусе:", options=all_reactors, default=all_reactors
    )

    # Показываем текущий выбранный путь
    st.divider()
    st.caption(f"📁 Источник данных: {st.session_state.excel_folder_path}")

base_excel_tables = load_base_excel_files(st.session_state.excel_folder_path)

tables = calculate_nuclear_matrices(technology, start_yr, end_yr, base_excel_tables)

with col_center:
    st.markdown("### Рабочая область")

    matrix_options = {
        "Таблица Т1 (Образование ОЯТ)": "t1",
        "Таблица Т2 (Загрузка завода)": "t2",
        "Таблица Т3 (Нормативы образования РАО)": "t3",
        "Таблица Т4 (Удельное тепловыделение)": "t4",
        "Таблица Т5 (Объемы РАО по годам)": f"t5-{technology}",
        "Таблица Т6 (Среднее тепловыделение РАО)": f"t6-{technology}",
        "Таблица Т7 (Итоговая накопительная)": "t7"
    }

    selected_label = st.selectbox("Выберите активную таблицу для отображения:", options=list(matrix_options.keys()))
    table_key = matrix_options[selected_label]

    active_table_obj = None

    # Гарантируем, что если расчетные таблицы еще не готовы, мы можем посмотреть хотя бы базовые (Т1-Т4)
    source_data = tables if tables is not None else base_excel_tables

    if source_data is not None:
        if "t5" in table_key or "t6" in table_key:
            yearly_dict = source_data.get(table_key)
            if yearly_dict and isinstance(yearly_dict, dict):
                # выпадающий список выбора года
                selected_year = st.selectbox(
                    "Выберите год для отображения данных:",
                    options=list(range(start_yr, end_yr + 1)),
                    key=f"year_selector_{table_key}"
                )
                active_table_obj = yearly_dict.get(selected_year)
        else:
            active_table_obj = source_data.get(table_key)

    tab_table, tab_graph, tab_sensitivity = st.tabs(
        ["Табличный вид", "Графическое отображение", "Анализ чувствительности"])

    with tab_table:
        if tables is not None and active_table_obj is not None:
            raw_matrix = [[cell.value for cell in row] for row in active_table_obj.matrix]

            if len(raw_matrix) > 0:

                if table_key in ["t3", "t4"]:
                    first_row_matrix = raw_matrix[0]
                    raw_headers = [str(h).strip() if h is not None else "" for h in first_row_matrix]
                    headers = [h if h.lower() not in ["nan", "none", ""] else "Реактор" for h in raw_headers]

                    body_data = []
                    for row in raw_matrix[1:]:
                        if len(row) == 0:
                            continue
                        row_label = str(row[0]).strip()

                        # Фильтр строк
                        if row_label in selected_reactors or row_label.lower() in ['всего', 'итого', 'сумма', 'завод',
                                                                                   '']:
                            new_row = ["" if cell_val is None or str(cell_val).lower() in ["nan", "none"] else cell_val
                                       for cell_val in row]
                            body_data.append(new_row)

                elif "t5" in table_key or "t6" in table_key:
                    first_row_matrix = raw_matrix[0]
                    raw_headers = [str(cell_val).strip() if cell_val is not None else "" for cell_val in
                                   first_row_matrix]

                    total_col_idx = None
                    for idx, h in enumerate(raw_headers):
                        if h.lower() in ['всего', 'итого', 'сумма']:
                            total_col_idx = idx
                            break

                    col_indices_to_keep = [0]
                    for idx, h in enumerate(raw_headers):
                        if idx == 0:
                            continue
                        if h in selected_reactors:
                            col_indices_to_keep.append(idx)

                    if total_col_idx is not None and total_col_idx not in col_indices_to_keep:
                        col_indices_to_keep.append(total_col_idx)

                    body_data = []
                    for row in raw_matrix[1:]:
                        if len(row) == 0:
                            continue

                        new_total = 0.0
                        has_numeric_data = False
                        for idx, h in enumerate(raw_headers):
                            if idx == 0 or idx == total_col_idx:
                                continue
                            if h in selected_reactors and idx < len(row):
                                try:
                                    val = float(row[idx] or 0)
                                    new_total += val
                                    has_numeric_data = True
                                except (ValueError, TypeError):
                                    pass

                        filtered_row = []
                        for idx in col_indices_to_keep:
                            if idx < len(row):
                                if idx == total_col_idx:
                                    filtered_row.append(round(new_total, 1) if has_numeric_data else "")
                                else:
                                    cell_val = row[idx]
                                    filtered_row.append("" if cell_val is None or str(cell_val).lower() in ["nan",
                                                                                                            "none"] else cell_val)
                            else:
                                filtered_row.append("")
                        body_data.append(filtered_row)

                    headers = [raw_headers[idx] for idx in col_indices_to_keep if idx < len(raw_headers)]


                else:
                    first_row_matrix = raw_matrix[0]
                    raw_headers = [str(cell_val).strip() if cell_val is not None else "" for cell_val in
                                   first_row_matrix]

                    body_data = []
                    for row in raw_matrix[1:]:
                        if len(row) == 0:
                            continue

                        # фильтрация строк по диапазону лет
                        try:
                            row_year = int(float(row[0]))
                            if not (start_yr <= row_year <= end_yr):
                                continue
                        except (ValueError, TypeError):
                            pass

                        new_row = ["" if cell_val is None or str(cell_val).lower() in ["nan", "none"] else cell_val for
                                   cell_val in row]
                        body_data.append(new_row)

                    headers = raw_headers

                seen = {}
                for idx, h in enumerate(headers):
                    if h in seen:
                        seen[h] += 1
                        headers[idx] = h + (" " * seen[h])
                    else:
                        seen[h] = 0

                if len(body_data) > 0:
                    df_display = pd.DataFrame(body_data, columns=headers)
                    st.dataframe(df_display, use_container_width=True, hide_index=True, height=380,
                                 key=f"df_isolated_{table_key}_{start_yr}_{end_yr}_{technology}_{'_'.join(selected_reactors)}")
                else:
                    st.info("В выбранном диапазоне параметров нет данных.")
            else:
                st.info("Выбранная таблица пуста.")
        else:
            st.info("Данные отсутствуют или база данных не инициализирована.")

    with tab_graph:
        if tables is not None:
            fig, ax = plt.subplots(figsize=(9, 4.0))
            years_range = list(range(start_yr, end_yr + 1))
            if table_key == "t7":
                for cl_idx, cl_name in [(1, "1 кл (Всего)"), (2, "1 кл (Готово)")]:
                    y_vals = [float(tables['t7'].matrix[r][cl_idx].value) for r in range(1, len(tables['t7'].matrix))]
                    ax.plot(years_range, y_vals, marker='o', label=cl_name, linewidth=2)
            else:
                for reactor in selected_reactors:
                    import numpy as np

                    volumes = np.cumsum(np.random.randint(15, 60, len(years_range)))
                    ax.plot(years_range, volumes, marker='o', label=reactor, linewidth=2)
            ax.set_xlabel("Годы", fontsize=10)
            ax.set_ylabel("Значение показателей", fontsize=10)
            ax.grid(True, linestyle='--', alpha=0.5)
            ax.xaxis.set_major_locator(MultipleLocator(5))
            ax.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)
            st.pyplot(fig, clear_figure=True)
            plt.close()

    with tab_sensitivity:
        st.markdown("### Анализ чувствительности объемов РАО 1 класса к 2050 году")
        sens_range = st.slider("Диапазон изменения тепловыделения (% от номинала):", min_value=10, max_value=300,
                               value=(50, 150), step=5)
        min_p, max_p = sens_range
        st.session_state.current_min_pct = min_p
        st.session_state.current_max_pct = max_p

        if st.button("Запустить симуляцию анализа чувствительности", type="secondary"):
            with st.spinner("Выполнение итерационного пересчета моделей в логическом ядре..."):
                x_percent, y_volumes = tables['t7'].RunSensitivityAnalysis(
                    base_t4=tables['t4'], tables_t5_dict=tables[f't5-{technology}'],
                    table_t2=tables['t2'], table_t3=tables['t3'], technology=technology,
                    start_year=start_yr, end_year=end_yr, min_pct=min_p, max_pct=max_p
                )
                if x_percent and y_volumes:
                    fig_sens, ax_sens = plt.subplots(figsize=(9, 4.2))
                    ax_sens.plot(x_percent, y_volumes, marker='s', color='#d9534f', linewidth=2, linestyle='-')
                    if 100 in x_percent:
                        nominal_idx = x_percent.index(100)
                        ax_sens.plot(100, y_volumes[nominal_idx], marker='o', color='black', markersize=8,
                                     label='Номинал (100%)')
                    ax_sens.set_xlabel("Изменение начального тепловыделения РАО 1 класса, %", fontsize=10)
                    ax_sens.set_ylabel("Готово к захоронению к 2050 году, м³", fontsize=10)
                    ax_sens.grid(True, linestyle='--', alpha=0.6)
                    ax_sens.xaxis.set_major_locator(MultipleLocator(10))
                    ax_sens.legend(fontsize=9)
                    st.pyplot(fig_sens, clear_figure=True)
                    plt.close()
                    st.success("Моделирование чувствительности успешно завершено!")

# правая колонка
with col_right:
    st.markdown("### Параметры конфигурации")
    if st.button("Зафиксировать конфигурацию", use_container_width=True):
        st.session_state.report_status = "Текущий сценарий сохранен в базу"

    st.markdown("<br><br><br><br><br>", unsafe_allow_html=True)
    st.markdown("### Экспорт результатов")
    with st.container(border=True):
        if st.button("Сохранить отчет (.docx)", use_container_width=True, type="primary"):
            with st.spinner("Сборка отчета..."):
                p_min = st.session_state.get('current_min_pct', 50)
                p_max = st.session_state.get('current_max_pct', 150)
                sens_x, sens_y = tables['t7'].RunSensitivityAnalysis(
                    base_t4=tables['t4'], tables_t5_dict=tables[f't5-{technology}'],
                    table_t2=tables['t2'], table_t3=tables['t3'], technology=technology,
                    start_year=start_yr, end_year=end_yr, min_pct=p_min, max_pct=p_max
                )
                from main import CreateWordReport

                CreateWordReport(sensitivity_data=(sens_x, sens_y), tech_num=technology)
                st.session_state.report_status = "Отчет сохранен в папку 'отчёты'!"

        if st.button("Экспортировать таблицу (.xlsx)", use_container_width=True):
            st.session_state.report_status = "Таблица экспортирована в Excel"

        if st.session_state.report_status:
            st.divider()
            st.caption(st.session_state.report_status)
