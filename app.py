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
if 'FuelClassHeatReady' not in st.session_state:
    st.session_state.FuelClassHeatReady = {"1 класс": 2.0, "2 класс": 0.5, "3 класс": 1.2}

if 'HalfLife' not in st.session_state:
    st.session_state.HalfLife = {"1 класс": 10.7, "2 класс": 50.0, "3 класс": 5.0}
if 'show_constants_popup' not in st.session_state:
    st.session_state.show_constants_popup = False
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
            "end_yr": 2100,
            "selected_reactors": ['ВВЭР-1000', 'ВВЭР-440', 'БН-600', 'БН-800', 'РБМК'],
            "current_min_pct": 50,
            "current_max_pct": 150
        }
    }
if 'active_plant_id' not in st.session_state:
    st.session_state.active_plant_id = 1
if 'next_plant_id' not in st.session_state:
    st.session_state.next_plant_id = 2
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
    clean_folder = folder_name.strip().replace("\\", "/")
    if clean_folder.startswith("excel/"):
        full_path = folder_name
    elif clean_folder == "excel":
        full_path = "excel"
    else:
        full_path = os.path.join("excel", folder_name)
    # существование папки
    if not os.path.exists(full_path):
        st.error(f"Ошибка: Папка '{full_path}' не найдена в директории 'excel'!")
        return None

    try:
        # Проверяем файлы t2, t3, t4 внутри папки завода
        for name in ['t2', 't3', 't4']:
            if not os.path.exists(os.path.join(full_path, f"{name}.xlsx")):
                st.error(f"В папке завода '{full_path}' отсутствует обязательный файл {name}.xlsx")
                return None

        plant_obj = ProcessingPlant(name=plant_name, folder_path=full_path)
        return plant_obj
    except Exception as e:
        st.error(f"Ошибка инициализации завода из папки '{full_path}': {e}")
        return None


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
                if f"plant_object_{active_id}" in st.session_state:
                    del st.session_state[f"plant_object_{active_id}"]

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
table_field_name = "t1"
col_left, col_center, col_right = st.columns([2.5, 6.5, 3.0], gap="large")
# Левая колонка
with col_left:
    st.markdown("### Панель управления")
    db_loaded = st.button("База данных", use_container_width=True)
    st.divider()
    st.markdown("**Выборка данных для расчета**")
    def on_technology_change():
        st.session_state.plants_data[active_id]["technology"] = st.session_state[f"tech_sel_{active_id}"]
        st.session_state[f"need_calc_{active_id}"] = True
    def on_slider_change():
        val = st.session_state[f"slider_{active_id}"]
        st.session_state.plants_data[active_id]["start_yr"] = val[0]
        st.session_state.plants_data[active_id]["end_yr"] = val[1]
        st.session_state[f"need_calc_{active_id}"] = True
    # Список технологий
    technology = st.selectbox(
        "Технология переработки:",
        options=[1, 2],
        index=current_config["technology"] - 1,
        key=f"tech_sel_{active_id}",
        on_change=on_technology_change  # Привязываем колбэк
    )
    # Слайдер для времиени (пока не используем)
    start_yr, end_yr = st.slider(
        "Временной диапазон (годы):",
        min_value=2030, max_value=2100,
        value=(current_config["start_yr"], current_config["end_yr"]),
        step=1,
        key=f"slider_{active_id}",
        on_change=on_slider_change
    )
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
    st.divider()
    if st.button("⚙️ Ввод констант", use_container_width=True, key="toggle_constants_btn"):
        st.session_state.show_constants_popup = not st.session_state.show_constants_popup
        st.rerun()


# Расчеты
plant_session_key = f"plant_object_{active_id}"
calc_trigger_key = f"need_calc_{active_id}"
# ТРиггер пересчета
if calc_trigger_key not in st.session_state:
    st.session_state[calc_trigger_key] = False
# Если завода для текущей вкладки нет в памяти то загружаем
if plant_session_key not in st.session_state:
    st.session_state[plant_session_key] = load_base_excel_files(current_config["folder_path"], current_config["name"])
    st.session_state[calc_trigger_key] = True
plant = st.session_state[plant_session_key]
if plant is not None:
    if st.session_state[calc_trigger_key]:
        plant.CalculateMatrices(technology=technology, start_year=start_yr, end_year=end_yr)
        # Протим бесконечной рекурсии
        st.session_state[calc_trigger_key] = False
plant_session_key = f"plant_object_{active_id}"
calc_trigger_key = f"need_calc_{active_id}"

if calc_trigger_key not in st.session_state:
    st.session_state[calc_trigger_key] = False

if plant_session_key not in st.session_state:
    st.session_state[plant_session_key] = load_base_excel_files(current_config["folder_path"], current_config["name"])
    st.session_state[calc_trigger_key] = True
plant = st.session_state[plant_session_key]
if plant is not None:
    if st.session_state[calc_trigger_key]:
        plant.CalculateMatrices(technology=technology, start_year=start_yr, end_year=end_yr)
        st.session_state[calc_trigger_key] = False
if st.session_state.show_constants_popup:
    # Перенаправляем вывод внутрь левой колонки за счет Streamlit контейнеров (внутри метода)
    InterfaceRenderer.RenderConstantsUI(
        active_id=active_id,
        technology=technology,
        start_yr=start_yr,
        end_yr=end_yr,
        plant=plant
    )
# центральная колонка
with col_center:
    if st.session_state.report_mode:
        InterfaceRenderer.RenderConstructorUI(
            start_yr=start_yr, end_yr=end_yr,
            current_config=current_config,
            active_id=active_id,
            technology=technology
        )

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
        storage_key = f"edited_df_{table_field_name}_{active_id}"
        for key in list(st.session_state.keys()):
            if "edited_df_" in key and key != storage_key:
                del st.session_state[key]
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
        tab_table, tab_graph, tab_sensitivity, tab_burial_graph, tab_capacity_analysis = st.tabs(
            [
                "Табличный вид",
                "Графическое отображение",
                "Анализ чувствительности",
                "Заполнение ПЗРО (График)",
                "Анализ мощностей"
            ]
        )
        with tab_table:
            InterfaceRenderer.RenderTableUI(
                active_table_obj=active_table_obj, table_field_name=table_field_name,
                active_id=active_id, start_yr=start_yr, end_yr=end_yr,
                technology=technology, selected_year=selected_year, plant=plant
            )
        with tab_graph:
            if plant is not None:
                years_range = list(range(start_yr, end_yr + 1))
                visualizer = NuclearDataVisualizer()
                if table_field_name in ["t3", "t4"]:
                    st.info(
                        "Графическое отображение для нормативов и удельного тепловыделения (Т3, Т4) не предусмотрено. Используйте Табличный вид.")
                elif not st.session_state.report_mode and table_field_name == "t7":
                    t7_field = getattr(plant, "t7", {})
                    active_t7_obj = t7_field.get(technology) if isinstance(t7_field, dict) else t7_field
                    if active_t7_obj is not None:
                        fig = visualizer.CreateT7PlotInCore(active_t7_obj, years_range)
                        st.pyplot(fig, clear_figure=True)
                        plt.close(fig)
                    else:
                        st.info("Таблица Т7 еще не рассчитана.")
                else:
                    # Базовый график реакторов из logic.py
                    fig = visualizer.CreateReactorDefaultPlot(selected_reactors, years_range)
                    st.pyplot(fig, clear_figure=True)
                    plt.close(fig)
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
        # Визуализация заполнения захоронений
        with tab_burial_graph:
            st.markdown("### График заполнения отраслевых ПЗРО накопительным итогом")
            st.caption(
                "Сплошные линии показывают фактический объем РАО в ПЗРО, пунктирные линии — максимальный проектный лимит объекта.")
            if plant is not None:
                visualizer = NuclearDataVisualizer()
                # Вызов готовых фигур из logic.py
                fig_pgzr, fig_ppzr = visualizer.CreateBurialPlotsInCore(st.session_state.plants_data, start_yr, end_yr)
                st.markdown("####Пункт глубинного захоронения РАО (ПГЗР)")
                st.pyplot(fig_pgzr, clear_figure=True)
                plt.close(fig_pgzr)
                st.divider()
                st.markdown("#### Пункт приповерхностного захоронения РАО (ППЗР)")
                st.pyplot(fig_ppzr, clear_figure=True)
                plt.close(fig_ppzr)
    with tab_capacity_analysis:
        st.markdown("### Системный анализ мощностей переработки и захоронения")
        st.caption(
            "Внимание: В связи с регламентом заполнения исходных данных, анализ ограничен периодом 2030–2050 гг.")

        if plant is not None:
            # Импортируем наши чистые расчетные функции из logic.py
            from logic import GetCapacityAnalysisData, GetBurialVsCapacityData
            st.markdown("#### 1. Баланс наработки ОЯТ и мощностей регенерации")
            st.caption(
                "Сравнение суммарного годового образования ОЯТ по всем реакторам (Т1) и фактической суммарной загрузки всех заводов (Т2).")

            cap_data = GetCapacityAnalysisData(st.session_state.plants_data, start_yr, end_yr)

            fig_cap1, ax_cap1 = plt.subplots(figsize=(9, 3.8))
            ax_cap1.plot(cap_data["years"], cap_data["t1_total"], linestyle='-', marker='o', color='#d9534f',
                         label="Образование ОЯТ (Т1, всего тонн)", linewidth=2.5)
            ax_cap1.plot(cap_data["years"], cap_data["t2_total_sum"], linestyle='--', marker='s', color='#2a6496',
                         label="Загрузка переработки (Т2, сумма заводов)", linewidth=2.0)

            ax_cap1.set_xlabel("Годы", fontsize=9)
            ax_cap1.set_ylabel("Масса ОЯТ, тонн ТМ", fontsize=9)
            ax_cap1.grid(True, linestyle='--', alpha=0.4)
            ax_cap1.xaxis.set_major_locator(MultipleLocator(5))
            ax_cap1.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=9)
            plt.tight_layout()

            st.pyplot(fig_cap1, clear_figure=True)
            plt.close(fig_cap1)

            st.divider()
            st.markdown("#### 2. Сопоставление объемов переработки и темпов кондиционирования РАО")
            st.caption(
                "Сравнение суммарной переработки ОЯТ на заводах (тонны) и результирующего потока кондиционированного РАО, готового к изоляции (кубометры, сумма 1–4 классов из Т7).")

            burial_vs_cap = GetBurialVsCapacityData(st.session_state.plants_data, start_yr, end_yr)

            fig_cap2, ax_cap2 = plt.subplots(figsize=(9, 3.8))

            # Используем две оси Y (двойная шкала), так как тонны и кубометры имеют разный масштаб!
            ax_cap2_left = ax_cap2
            ax_cap2_right = ax_cap2.twinx()

            # Левая шкала: Тонны ОЯТ
            line1 = ax_cap2_left.plot(burial_vs_cap["years"], burial_vs_cap["total_t2"], linestyle='-', marker='^',
                                      color='#2a6496', label="Переработка ОЯТ (Т2, тонн)", linewidth=2.0)
            ax_cap2_left.set_ylabel("Загрузка заводов, тонн ТМ", color='#2a6496', fontsize=9)
            ax_cap2_left.tick_params(axis='y', labelcolor='#2a6496')

            # Правая шкала: Кубометры готового РАО
            line2 = ax_cap2_right.plot(burial_vs_cap["years"], burial_vs_cap["total_burial_ready"], linestyle='-',
                                       marker='s', color='#2b702b', label="Готово к захоронению РАО (Т7, м³)",
                                       linewidth=2.0)
            ax_cap2_right.set_ylabel("Поток готового РАО, м³", color='#2b702b', fontsize=9)
            ax_cap2_right.tick_params(axis='y', labelcolor='#2b702b')

            ax_cap2_left.set_xlabel("Годы", fontsize=9)
            ax_cap2_left.grid(True, linestyle='--', alpha=0.4)
            ax_cap2_left.xaxis.set_major_locator(MultipleLocator(5))

            # Объединяем легенды двух осей в одну общую
            lines = line1 + line2
            labels = [l.get_label() for l in lines]
            ax_cap2_left.legend(lines, labels, loc='upper left', bbox_to_anchor=(1.1, 1), fontsize=9)
            plt.tight_layout()

            st.pyplot(fig_cap2, clear_figure=True)
            plt.close(fig_cap2)

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