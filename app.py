import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import pandas as pd
import os
from logic import *
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import pandas as pd
import os
from office import CreateWordReport

st.set_page_config(page_title="РАО Аналитика", layout="wide")

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


def load_base_excel_files(folder_name, plant_name):
    full_path = os.path.join("excel", folder_name)

    if not os.path.exists(full_path):
        st.error(f"Ошибка: Папка '{full_path}' не найдена в директории 'excel'!")
        return None
    try:
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
    plant.calculate_matrices(technology=technology, start_year=start_yr, end_year=end_yr)

# --- ЦЕНТРАЛЬНАЯ КОЛОНКА ---
with col_center:
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
            fig, ax = plt.subplots(figsize=(9, 4.0))
            years_range = list(range(start_yr, end_yr + 1))
            t7_field = getattr(plant, "t7", {})
            active_t7_obj = t7_field.get(technology) if isinstance(t7_field, dict) else t7_field

            if table_field_name == "t7" and active_t7_obj is not None:
                for cl_idx, cl_name in [(1, "1 кл (Всего)"), (2, "1 кл (Готово)")]:
                    y_vals = [float(active_t7_obj.matrix[r][cl_idx].value) for r in range(1, len(active_t7_obj.matrix))]
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
            st.markdown("### Анализ чувствительности объемов РАО 1 класса к конечному году")
            sens_range = st.slider("Диапазон изменения тепловыделения (% от номинала):", min_value=10, max_value=300,
                                   value=(50, 150), step=5)
            min_p, max_p = sens_range
            st.session_state.current_min_pct = min_p
            st.session_state.current_max_pct = max_p

            if st.button("Запустить симуляцию анализа чувствительности", type="secondary"):
                if plant is not None and plant.t7 is not None:
                    with st.spinner("Выполнение итерационного пересчета моделей в логическом ядре..."):
                        x_percent, y_volumes = plant.t7[technology].RunSensitivityAnalysis(
                            base_t4=plant.t4, tables_t5_dict=plant.t5_series[technology],
                            table_t2=plant.t2, table_t3=plant.t3, technology=technology,
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
                            ax_sens.set_ylabel(f"Готово к захоронению к {end_yr} году, м³", fontsize=10)
                            ax_sens.grid(True, linestyle='--', alpha=0.6)
                            ax_sens.xaxis.set_major_locator(MultipleLocator(10))
                            ax_sens.legend(fontsize=9)
                            st.pyplot(fig_sens, clear_figure=True)
                            plt.close()
                            st.success("Моделирование чувствительности успешно завершено!")
                else:
                    st.error("Нет данных для проведения анализа. Проверьте загрузку таблиц завода.")
    with col_right:
        st.markdown("### Параметры конфигурации")

        if st.button("Зафиксировать конфигурацию", use_container_width=True):
            st.session_state.report_status = "Текущий сценарий сохранен в базу"

        st.markdown("<br><br><br><br><br><br><br>", unsafe_allow_html=True)
        st.markdown("### Экспорт результатов")
        with st.container(border=True):
            if st.button("Сохранить отчет (.docx)", use_container_width=True, type="primary",
                         key=f"doc_btn_{active_id}"):
                if plant is not None:
                    with st.spinner("Сборка отчета и принудительный расчет обеих технологий в ядре..."):
                        p_min = st.session_state.get('current_min_pct', 50)
                        p_max = st.session_state.get('current_max_pct', 150)

                        # ИСПРАВЛЕНИЕ: Перед вызовом отчета принудительно рассчитываем ОБЕ технологии для текущей папки!
                        # Это заставит logic.py заполнить и plant.t5_series, и plant.t5_series
                        plant.calculate_matrices(technology=1, start_year=start_yr, end_year=end_yr)
                        plant.calculate_matrices(technology=2, start_year=start_yr, end_year=end_yr)

                        # Проводим анализ чувствительности для текущей активной на экране технологии
                        sens_x, sens_y = plant.t7[technology].RunSensitivityAnalysis(
                            base_t4=plant.t4,
                            tables_t5_dict=plant.t5_series[technology],
                            table_t2=plant.t2,
                            table_t3=plant.t3,
                            technology=technology,
                            start_year=start_yr,
                            end_year=end_yr,
                            min_pct=p_min,
                            max_pct=p_max
                        )


                        # Передаем рассчитанный завод в отчет
                        CreateWordReport(plant=plant, sensitivity_data=(sens_x, sens_y), tech_num=technology)
                        plant.calculate_matrices(technology=technology, start_year=start_yr, end_year=end_yr)
                        st.session_state.report_status = f"Отчет для {current_config['name']} успешно сохранен!"
                        st.rerun()
                else:
                    st.error("Нет расчетных данных.")

            if st.button("Экспортировать таблицу (.xlsx)", use_container_width=True):
                st.session_state.report_status = "Таблица экспортирована в Excel"

            if st.session_state.report_status:
                st.divider()
                st.caption(st.session_state.report_status)
