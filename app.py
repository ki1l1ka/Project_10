from logic import *
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import pandas as pd
import os
from office import ReportDocument
if 'report_mode' not in st.session_state:
    st.session_state.report_mode = False
if 'constructor_queue' not in st.session_state:
    st.session_state.constructor_queue = []
# Настройка страницы
st.set_page_config(page_title="РАО Аналитика", layout="wide")

st.markdown("""
    <style>
    # Приведение к дескопному виду
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    # Фикс колонок
    [data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: stretch !important;
        width: 100% !important;
    }
    [data-testid="stHorizontalBlock"] > div:nth-child(1) {
        min-width: 270px !important; 
    }
    [data-testid="stHorizontalBlock"] > div:nth-child(2) {
        flex-grow: 2 !important;    
    }
    [data-testid="stHorizontalBlock"] > div:nth-child(3) {
        min-width: 290px !important;
    }

   # Внутренние поля
    .block-container {
        padding-top: 0.5rem;
        padding-left: 1.5rem;
        padding-right: 1.5rem;
    }

    # Вкладки
    div.stButton > button {
        border-radius: 4px !important;
        font-weight: 500 !important;
        transition: all 0.2s ease-in-out !important;
    }

    /* Эффект при наведении курсора на кнопки-вкладки заводов */
    div.stButton > button:hover {
        border-color: var(--primary-color) !important;
        box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.15) !important;
    }

    # Таблицы и сетка
    [data-testid="stDataFrame"] {
        border: 1px solid rgba(128, 128, 128, 0.2) !important;
        border-radius: 6px !important;
        overflow: hidden !important;
    }

    # Инф. Блоки
    [data-testid="stBlock"] {
        border-radius: 6px !important;
    }
    </style>
    """, unsafe_allow_html=True)


def load_base_excel_files(folder_name, plant_name):
    # Полный путь к папке завода
    full_path = os.path.join("excel", folder_name)
    if not os.path.exists(os.path.join("excel", "t1.xlsx")):
        st.error("Ошибка: Общая таблица t1.xlsx отсутствует в корне папки 'excel'!")
        return None

    if not os.path.exists(full_path):
        st.error(f"Ошибка: Папка '{full_path}' не найдена в директории 'excel'!")
        return None
    try:
        # Проверяем файлы t2, t3, t4 внутри индивидуальной папки
        for name in ['t2', 't3', 't4']:
            if not os.path.exists(os.path.join(full_path, f"{name}.xlsx")):
                st.error(f"В папке завода 'excel/{folder_name}' отсутствует обязательный файл {name}.xlsx")
                return None

        # Создаем объект завода
        plant_obj = ProcessingPlant(name=plant_name, folder_path=full_path)
        return plant_obj
    except Exception as e:
        st.error(f"Ошибка инициализации завода из папки '{full_path}': {e}")
        return None


if 'report_status' not in st.session_state:
    st.session_state.report_status = ""

# путь
if 'plants_data' not in st.session_state:
    st.session_state.plants_data = {
        1: {
            "name": "Завод 1",
            "folder_path": "Завод 1",
            "technology": 1,
            "start_yr": 2030,
            "end_yr": 2050,
            "selected_reactors": ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК'],
            "current_min_pct": 50,
            "current_max_pct": 150
        }
    }

if 'active_plant_id' not in st.session_state:
    st.session_state.active_plant_id = 1
if 'next_plant_id' not in st.session_state:
    st.session_state.next_plant_id = 2

with st.container():
    num_plants = len(st.session_state.plants_data)
    top_cols = st.columns([1.5, 3.5] + [1.0] * num_plants + [1.0])

    with top_cols[0]:
        if st.button("Добавить завод", use_container_width=True):
            new_id = st.session_state.next_plant_id
            st.session_state.plants_data[new_id] = {
                "name": f"Завод {new_id}",
                "folder_path": "excel",
                "technology": 1,
                "start_yr": 2030,
                "end_yr": 2050,
                "selected_reactors": ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК'],
                "current_min_pct": 50,
                "current_max_pct": 150
            }
            st.session_state.active_plant_id = new_id
            st.session_state.next_plant_id += 1
            st.rerun()

    active_id = st.session_state.active_plant_id
    current_config = st.session_state.plants_data[active_id]

    with top_cols[1]:
        input_folder = st.text_input(
            "Название папки завода:",
            value=current_config["folder_path"],
            placeholder="Например: Завод 1",
            label_visibility="collapsed",
            key=f"path_input_{active_id}"
        )
        if input_folder != current_config["folder_path"]:
            check_path = os.path.join("excel", input_folder)
            if os.path.exists(check_path):
                st.session_state.plants_data[active_id]["folder_path"] = input_folder
                st.toast(f"Для '{current_config['name']}' установлена папка: excel/{input_folder}")
                st.rerun()
            else:
                st.error(f"Папка 'excel/{input_folder}' не найдена!")

    col_idx = 2
    for p_id, p_info in st.session_state.plants_data.items():
        with top_cols[col_idx]:
            btn_type = "primary" if p_id == active_id else "secondary"
            if st.button(p_info["name"], key=f"tab_btn_{p_id}", use_container_width=True, type=btn_type):
                st.session_state.active_plant_id = p_id
                st.rerun()
        col_idx += 1

st.divider()
active_table_obj = None
selected_year = "all"
table_field_name = "t1"  # По умолчанию выставляем базовую таблицу Т1

col_left, col_center, col_right = st.columns([2.5, 6.5, 3.0], gap="large")

with col_left:
    st.markdown("### Панель управления")
    db_loaded = st.button("База данных", use_container_width=True)

    st.divider()
    st.markdown("**Выборка данных для расчета**")

    technology = st.selectbox(
        "Технология переработки:",
        options=[1, 2],
        index=current_config["technology"] - 1,
        key=f"tech_sel_{active_id}"
    )
    st.session_state.plants_data[active_id]["technology"] = technology

    start_yr, end_yr = st.slider(
        "Временной диапазон (годы):",
        min_value=2030, max_value=2050,
        value=(current_config["start_yr"], current_config["end_yr"]),
        step=1,
        key=f"slider_{active_id}"
    )
    st.session_state.plants_data[active_id]["start_yr"] = start_yr
    st.session_state.plants_data[active_id]["end_yr"] = end_yr

    st.divider()
    st.markdown("**Инструменты и реакторы**")
    all_reactors = ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК']

    selected_reactors = st.multiselect(
        "Установки в фокусе:",
        options=all_reactors,
        default=current_config["selected_reactors"],
        key=f"reactors_{active_id}"
    )
    st.session_state.plants_data[active_id]["selected_reactors"] = selected_reactors

    st.divider()
    st.caption(f"Источник данных: {current_config['folder_path']}")
plant = load_base_excel_files(current_config["folder_path"], current_config["name"])

if plant is not None:
    plant.CalculateMatrices(technology=technology, start_year=start_yr, end_year=end_yr)

# централяная колонка
with col_center:
        if st.session_state.report_mode:
            st.markdown("### 🛠️ Конструктор отчета Word (.docx)")
            st.caption("Выберите элементы программы, настройте их масштаб и порядок для записи в итоговый документ.")

            from logic import ReportConstructor

            constructor = ReportConstructor()
            # Получаем полный список всех доступных таблиц и графиков по всем заводам
            available_items = constructor.GetAvailableElements(st.session_state.plants_data, start_yr, end_yr)
            with st.container(border=True):
                st.markdown("**Шаг 1. Добавить элемент в документ**")
                item_options = {item["label"] + (f" ({item['group']})" if item['plant_name'] else ""): item for item in
                                available_items}
                selected_item_label = st.selectbox("Выберите таблицу или график для вставки:",
                                                   options=list(item_options.keys()))

                # Ползунок масштабирования (пока в дюймах)
                element_size = st.slider("Ширина элемента в документе (дюймы):", min_value=2.0, max_value=8.0,
                                         value=5.5, step=0.5)

                if st.button("➕ Добавить выбранный элемент в очередь отчета", use_container_width=True):
                    chosen_meta = item_options[selected_item_label]
                    # Записываем элемент в очередь
                    st.session_state.constructor_queue.append({
                        "key": chosen_meta["key"],
                        "type": chosen_meta["type"],
                        "label": chosen_meta["label"],
                        "field_name": chosen_meta["field_name"],
                        "plant_name": chosen_meta["plant_name"],
                        "tech": chosen_meta.get("tech"),
                        "year": chosen_meta.get("year"),
                        "size": element_size
                    })
                    st.toast("Элемент добавлен в очередь отчета!")
                    st.rerun()

            st.markdown("**Шаг 2. Текущая структура документа (Очередь сборки)**")
            if not st.session_state.constructor_queue:
                st.info("Очередь пуста. Добавьте элементы выше.")
            else:
                # Выводим очередь в виде списка
                for idx, q_item in enumerate(st.session_state.constructor_queue):
                    q_col_text, q_col_size, q_col_del = st.columns([6.0, 2.0, 1.0])
                    with q_col_text:
                        prefix = "таблица" if q_item["type"] == "table" else "график"
                        plant_prefix = f"[{q_item['plant_name']}] " if q_item["plant_name"] else "[Глобальный] "
                        st.write(f"{idx + 1}. {prefix} {plant_prefix}{q_item['label']}")
                    with q_col_size:
                        st.caption(f"Ширина: {q_item['size']} дюйм.")
                    with q_col_del:
                        if st.button("удалить", key=f"del_q_{idx}"):
                            st.session_state.constructor_queue.pop(idx)
                            st.rerun()

                st.divider()
                btn_build_col, btn_cancel_col = st.columns(2)

                with btn_build_col:
                    # финальная сборка
                    if st.button("Сгенерировать и сохранить отчет Word", use_container_width=True, type="primary"):
                        with st.spinner("Идет генерация документа конструктором..."):
                            from office import ReportDocument

                            doc_report = ReportDocument(title="Аналитический комплекс: Пользовательский отчет")

                            for elem in st.session_state.constructor_queue:
                                # Имя элемента как заголовок в Word
                                plant_title = f"Завод: {elem['plant_name']} | " if elem['plant_name'] else ""
                                doc_report.AddHeading(f"{plant_title}{elem['label']}", level=1)
                                current_plant_obj = None
                                if elem['plant_name']:
                                    for p_id, p_info in st.session_state.plants_data.items():
                                        if p_info["name"] == elem['plant_name']:
                                            current_plant_obj = load_base_excel_files(p_info["folder_path"],
                                                                                      p_info["name"])
                                            if current_plant_obj:
                                                current_plant_obj.CalculateMatrices(p_info["technology"],
                                                                                    p_info["start_yr"],
                                                                                    p_info["end_yr"])
                                            break

                                # Сборка таблицы
                                if elem["type"] == "table":
                                    target_table = None
                                    if elem["key"] == "global_t1" and current_plant_obj:
                                        target_table = current_plant_obj.t1
                                    elif current_plant_obj:
                                        if "series" in elem["field_name"]:
                                            series_dict = getattr(current_plant_obj, elem["field_name"]).get(
                                                elem["tech"], {})
                                            target_table = series_dict.get(elem["year"])
                                        elif elem["field_name"] == "t7":
                                            target_table = current_plant_obj.t7.get(elem["tech"])
                                        else:
                                            target_table = getattr(current_plant_obj, elem["field_name"])

                                    if target_table:
                                        doc_report.AddParagraph(getattr(target_table, 'description', ''))
                                        doc_report.AddTable(target_table)
                                        doc_report.AddPageBreak()

                                # Сбокаграфика
                                elif elem["type"] == "graph" and current_plant_obj:
                                    fig, ax = plt.subplots(figsize=(10, 5))
                                    if elem["field_name"] == "graph_t7" and current_plant_obj.t7.get(
                                            current_config["technology"]):
                                        active_t7 = current_plant_obj.t7[current_config["technology"]]

                                        visualizer = NuclearDataVisualizer()
                                        t7_data = visualizer.GetT7PlotData(active_t7_obj=active_t7)

                                        fig, ax = plt.subplots(figsize=(10, 5))
                                        colors = {
                                            "1 класс (Всего)": "#d9534f", "1 класс (Готово)": "#942a27",
                                            "2 класс (Всего)": "#f0ad4e", "2 класс (Готово)": "#b57d28",
                                            "3 класс (Всего)": "#5cb85c", "3 класс (Готово)": "#2b702b"
                                        }

                                        for label_name, y_values in t7_data.items():
                                            line_style = '--' if 'Всего' in label_name else '-'
                                            ax.plot(years_range, y_values, linestyle=line_style, marker='o',
                                                    color=colors[label_name], label=label_name)

                                        ax.grid(True, linestyle='--', alpha=0.5)
                                        ax.legend()
                                        doc_report.AddPlot(fig, title=elem["label"], width_inches=elem["size"])
                                        doc_report.AddPageBreak()

                                    elif elem["field_name"] == "graph_sens" and current_plant_obj.t7.get(
                                            current_config["technology"]):
                                        p_min_c = st.session_state.get('current_min_pct', 50)
                                        p_max_c = st.session_state.get('current_max_pct', 150)
                                        x_p, y_v = current_plant_obj.RunSensitivityAnalysis(
                                            current_config["technology"], start_yr, end_yr, p_min_c, p_max_c)
                                        ax.plot(x_p, y_v, marker='s', color='#d9534f', linewidth=2)

                                    ax.set_xlabel("Показатели")
                                    ax.grid(True, linestyle='--', alpha=0.5)
                                    ax.legend()
                                    doc_report.AddPlot(fig, title=elem["label"], width_inches=elem["size"])
                                    doc_report.AddPageBreak()

                            # Сохраняем и выводим имя созданного файла
                            saved_file = doc_report.Save()
                            st.session_state.report_status = f"Пользовательский отчет '{os.path.basename(saved_file)}' успешно собран конструктором!"
                            st.session_state.constructor_queue = []  # Очищаем очередь
                            st.session_state.report_mode = False  # Выключаем конструктор
                            st.rerun()

                with btn_cancel_col:
                    if st.button("Закрыть конструктор (Вернуться к таблицам)", use_container_width=True):
                        st.session_state.report_mode = False
                        st.rerun()

        else:
            st.markdown("### Рабочая область")

            matrix_options = {
                "Таблица Т1 (Образование ОЯТ)": "t1",
                "Таблица Т2 (Загрузка завода)": "t2",
                "Таблица Т3 (Нормативы образования РАО)": "t3",
                "Таблица Т4 (Удельное тепловыделение)": "t4",
                "Таблица Т5 (Объемы РАО по годам)": "t5_series",
                "Таблица Т6 (Среднее тепловыделение РАО)": "t6_series",
                "Таблица Т7 (Итоговая накопительная)": "t7"
            }

            selected_label = st.selectbox("Выберите активную таблицу для отображения:", options=list(matrix_options.keys()))
            table_field_name = matrix_options[selected_label]
            active_table_obj = None
            selected_year = "all"

            if plant is not None:
                if "series" in table_field_name:
                    available_years = list(range(start_yr, end_yr + 1))
                    selected_year = st.selectbox(
                        "Выберите год для отображения данных:",
                        options=available_years,
                        key=f"year_selector_{table_field_name}_{active_id}"
                    )

                    plant_series_dict = getattr(plant, table_field_name)

                    if not isinstance(plant_series_dict, dict):
                        plant_series_dict = {1: {}, 2: {}}
                        setattr(plant, table_field_name, plant_series_dict)

                    if technology not in plant_series_dict or not plant_series_dict[technology]:
                        plant.calculate_matrices(technology=technology, start_year=start_yr, end_year=end_yr)
                        plant_series_dict = getattr(plant, table_field_name)

                    tech_dict = plant_series_dict.get(technology, {})

                    active_table_obj = tech_dict.get(selected_year)
                else:
                    active_table_obj = getattr(plant, table_field_name)
                    if table_field_name == "t7" and active_table_obj is not None:
                        if isinstance(active_table_obj, dict):
                            if technology not in active_table_obj or active_table_obj[technology] is None:
                                plant.calculate_matrices(technology=technology, start_year=start_yr, end_year=end_yr)
                                active_table_obj = getattr(plant, "t7")
                            active_table_obj = active_table_obj.get(technology)
                        else:
                            pass

        tab_table, tab_graph, tab_sensitivity = st.tabs(
            ["Табличный вид", "Графическое отображение", "Анализ чувствительности"]
        )
        with tab_table:
            if active_table_obj is not None:
                raw_matrix = [[cell.value for cell in row] for row in active_table_obj.matrix]

                if len(raw_matrix) > 0:
                    max_cols = max(len(row) for row in raw_matrix)
                    headers = []
                    first_row = raw_matrix[0]
                    for cell_val in first_row:
                        if cell_val is None:
                            val_str = ""
                        elif isinstance(cell_val, (list, tuple)):
                            val_str = " ".join([str(v).strip() for v in cell_val if v is not None])
                        else:
                            val_str = str(cell_val).strip()

                        if val_str.lower() == "nan" or val_str == "None":
                            val_str = ""
                        headers.append(val_str)
                    if len(headers) < max_cols:
                        headers.extend([""] * (max_cols - len(headers)))
                    elif len(headers) > max_cols:
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
                        if len(row) == 0:
                            continue
                        new_row = []
                        for cell_val in row:
                            if cell_val is None or str(cell_val).lower() == "nan" or str(cell_val) == "None":
                                new_row.append("")
                            else:
                                new_row.append(cell_val)

                        if len(new_row) < max_cols:
                            new_row.extend([""] * (max_cols - len(new_row)))
                        elif len(new_row) > max_cols:
                            new_row = new_row[:max_cols]
                        body_data.append(new_row)
                    df_display = pd.DataFrame(body_data, columns=headers)
                    st.dataframe(
                        df_display, use_container_width=True, hide_index=True, height=380,
                        key=f"df_{table_field_name}_{start_yr}_{end_yr}_{technology}_{selected_year}"
                    )
                else:
                    st.info("Выбранная таблица пуста.")
            else:
                st.info("Данные отсутствуют или база данных не инициализирована.")

        with tab_graph:
            if plant is not None:
                years_range = list(range(start_yr, end_yr + 1))
                if not st.session_state.report_mode and table_field_name == "t7":
                    t7_field = getattr(plant, "t7", {})
                    active_t7_obj = t7_field.get(technology) if isinstance(t7_field, dict) else t7_field

                    if active_t7_obj is not None:
                        visualizer = NuclearDataVisualizer()
                        t7_data = visualizer.GetT7PlotData(active_t7_obj=active_t7_obj)
                        fig, ax = plt.subplots(figsize=(9, 4.2))

                        # цвета для отрисовки массивов
                        colors = {
                            "1 класс (Всего)": "#d9534f", "1 класс (Готово)": "#942a27",
                            "2 класс (Всего)": "#f0ad4e", "2 класс (Готово)": "#b57d28",
                            "3 класс (Всего)": "#5cb85c", "3 класс (Готово)": "#2b702b"
                        }

                        # Отрисовываем 6 линий на холсте в цикле
                        for label_name, y_values in t7_data.items():
                            if len(years_range) == len(y_values):
                                line_style = '--' if 'Всего' in label_name else '-'
                                marker_style = 'o' if 'Всего' in label_name else 's'
                                marker_size = 3 if 'Всего' in label_name else 4
                                line_width = 1.5 if 'Всего' in label_name else 2.5

                                ax.plot(years_range, y_values, linestyle=line_style, marker=marker_style,
                                        markersize=marker_size, color=colors[label_name],
                                        label=label_name, linewidth=line_width)

                        # Стилизация осей
                        ax.set_xlabel("Годы", fontsize=10)
                        ax.set_ylabel("Объем РАО, м³", fontsize=10)
                        ax.grid(True, linestyle='--', alpha=0.5)
                        ax.xaxis.set_major_locator(MultipleLocator(5))
                        ax.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)

                        st.pyplot(fig, clear_figure=True)
                        plt.close()
                    else:
                        st.info("Таблица Т7 еще не рассчитана.")

                else:
                    fig, ax = plt.subplots(figsize=(9, 4.0))
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
                st.markdown("### Анализ чувствительности объемов РАО 1 класса к конечному году")
                sens_range = st.slider("Диапазон изменения тепловыделения (% от номинала):", min_value=10, max_value=300,
                                       value=(50, 150), step=5)
                min_p, max_p = sens_range
                st.session_state.current_min_pct = min_p
                st.session_state.current_max_pct = max_p

                if st.button("Запустить симуляцию анализа чувствительности", type="secondary"):
                    if plant is not None and plant.t7 is not None:
                        with st.spinner("Выполнение итерационного пересчета моделей в логическом ядре..."):
                            x_percent, y_volumes = plant.RunSensitivityAnalysis(
                                technology=technology,
                                start_year=start_yr,
                                end_year=end_yr,
                                min_pct=min_p,
                                max_pct=max_p
                            )
                            if x_percent and y_volumes:
                                fig_sens, ax_sens = plt.subplots(figsize=(9, 4.2))
                                ax_sens.plot(x_percent, y_volumes, marker='s', color='#d9534f', linewidth=2, linestyle='-')
                                if 100 in x_percent:
                                    nominal_idx = x_percent.index(100)
                                    ax_sens.plot(100, y_volumes[nominal_idx], marker='o', color='black', markersize=8,
                                                 label='Номинал (100%)')
                                ax_sens.set_xlabel("Изменение начального тепловыделения РАО 1 класса, %", fontsize=10)
                                ax_sens.set_ylabel(f"Готово к захоронению к {end_yr} году, м³", fontsize=10)
                                ax_sens.grid(True, linestyle='--', alpha=0.6)
                                ax_sens.xaxis.set_major_locator(MultipleLocator(10))
                                ax_sens.legend(fontsize=9)
                                st.pyplot(fig_sens, clear_figure=True)
                                plt.close()
                                st.success("Моделирование чувствительности успешно завершено!")
                    else:
                        st.error("Нет данных для проведения анализа. Проверьте загрузку таблиц завода.")
# Правая колонка
with col_right:
    st.markdown("### Параметры конфигурации")
    if st.button("Зафиксировать конфигурацию", use_container_width=True, key=f"fix_btn_{active_id}"):
        st.session_state.report_status = f"Сценарий для {current_config['name']} зафиксирован"

    st.markdown("<br><br><br><br><br><br><br>", unsafe_allow_html=True)
    st.markdown("### Экспорт результатов")
    with st.container(border=True):
        if st.button("Открыть Конструктор отчетов", use_container_width=True, type="primary"):
            st.session_state.report_mode = True
            st.rerun()

        if st.button("Экспортировать таблицу (.xlsx)", use_container_width=True, key=f"xlsx_btn_{active_id}"):
            if active_table_obj is not None:
                from logic import ExportTableToExcel

                year_suffix = f"_{selected_year}" if selected_year != "all" else ""
                exported_path = ExportTableToExcel(table_obj=active_table_obj,
                                                   label_name=f"{table_field_name}{year_suffix}")
                if exported_path:
                    st.session_state.report_status = f"Файл '{os.path.basename(exported_path)}' сохранен!"
                    st.rerun()
            else:
                st.error("Нет активной таблицы для экспорта.")

        if st.session_state.report_status:
            st.divider()
            st.caption(st.session_state.report_status)

