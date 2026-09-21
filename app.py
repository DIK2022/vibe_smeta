# app.py
# Интерактивная демонстрация. Бонус к ТЗ.
# Запуск: uv run streamlit run app.py

import streamlit as st
import plotly.graph_objects as go
import numpy as np

from core import (
    ROOMS, FURNITURE, COLOR_MAP, WALL_HEIGHT,
    build_all, compute_volumes, export_excel_bytes,
    compare_sources, get_open_questions, NAMING_CONFLICTS,
)

st.set_page_config(page_title="PDF → 3D → Смета", layout="wide")


def plot_3d(verts, faces, colors, include_ceiling=False):
    fig = go.Figure()
    order = ["floor_wood", "floor_tile", "ceiling", "wall", "furniture"]
    if not include_ceiling:
        order = [c for c in order if c != "ceiling"]

    for cname in order:
        mask = colors == cname
        if not mask.any():
            continue
        f_sub = faces[np.where(mask[faces[:, 0]])[0]]
        fig.add_trace(go.Mesh3d(
            x=verts[:, 0], y=verts[:, 1], z=verts[:, 2],
            i=f_sub[:, 0], j=f_sub[:, 1], k=f_sub[:, 2],
            color=COLOR_MAP[cname],
            opacity=0.35 if cname == "ceiling" else 1.0,
            name=cname, flatshading=True,
            lighting=dict(ambient=0.6, diffuse=0.9, specular=0.1, roughness=0.9),
        ))

    fig.update_layout(
        scene=dict(
            xaxis_title="X, м", yaxis_title="Y, м", zaxis_title="Z, м",
            aspectmode="data",
            camera=dict(eye=dict(x=1.2, y=-1.3, z=1.4), up=dict(x=0, y=0, z=1)),
            bgcolor="#f0f2f6",
        ),
        height=700, margin=dict(l=0, r=0, t=10, b=0),
        legend=dict(orientation="h", y=1.02, x=0),
    )
    return fig


def style_compare(df):
    def row_style(row):
        status = str(row["Статус"])
        if row["Помещение"] == "ИТОГО":
            return ["background-color: #dddddd; font-weight: bold"] * len(row)
        if "конфликт назначения" in status:
            return ["background-color: #ffeb9c"] * len(row)
        if "РАСХОЖДЕНИЕ" in status:
            return ["background-color: #ffc7ce"] * len(row)
        if status == "OK":
            return ["background-color: #c6efce"] * len(row)
        if "нет на чертеже" in status:
            return ["background-color: #ffd8a8"] * len(row)
        return [""] * len(row)
    return df.style.apply(row_style, axis=1)


# ============================================================
# UI
# ============================================================
st.title("🏗️ PDF-планировка → 3D-модель → Excel-смета")
st.caption("Прототип для вакансии «Вайбкодер / AI-разработчик»")

df_cmp = compare_sources()
n_warn = df_cmp["Статус"].astype(str).str.contains("РАСХОЖДЕНИЕ").sum()
n_conflict = len(NAMING_CONFLICTS)

if n_warn or n_conflict:
    parts = []
    if n_warn:
        parts.append(f"**{n_warn}** расхождений площадей")
    if n_conflict:
        parts.append(f"**{n_conflict}** конфликт назначения")
    st.warning(
        "⚠️ Обнаружены проблемы в исходных данных: " + ", ".join(parts) +
        ". Подробнее — вкладки «🔍 Сверка источников» и «❗ Открытые вопросы»."
    )
else:
    st.success("✅ Экспликация и чертёж согласованы.")

st.sidebar.header("Источник площадей")
source = st.sidebar.radio(
    "Что использовать в расчёте:",
    options=["explication", "drawing", "geometry"],
    format_func=lambda x: {
        "explication": "Экспликация (официальный документ)",
        "drawing":     "Чертёж (кружки на плане)",
        "geometry":    "Геометрия 3D-модели",
    }[x],
    index=0,
)
st.sidebar.caption(
    "По умолчанию — экспликация. Это официальный документ, "
    "по которому работает смета."
)

DF_VOL = compute_volumes(source=source)

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📐 3D-модель", "🔍 Сверка источников", "❗ Открытые вопросы",
    "📊 Объёмы", "📄 Excel", "ℹ️ О решении",
])

with tab1:
    st.subheader("Интерактивная 3D-модель квартиры")
    show_ceiling = st.checkbox("Показать потолки (полупрозрачные)", value=False)
    VERTS, FACES, COLORS = build_all(include_ceiling=show_ceiling)
    fig = plot_3d(VERTS, FACES, COLORS, include_ceiling=show_ceiling)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "Бежевый пол — паркет, серо-голубой — плитка. "
        "Стены — серые, мебель — зелёная."
    )

with tab2:
    st.subheader("Сверка: Экспликация vs Чертёж")
    st.markdown(
        "На обмерном плане площади в экспликации и кружки на чертеже "
        "расходятся. Это свойство исходных данных, а не ошибка распознавания."
    )

    st.dataframe(
        style_compare(df_cmp),
        use_container_width=True,
        hide_index=True,
    )

    st.info(
        "🟩 **зелёный** — расхождение ≤ 5 %. "
        "🟥 **красный** — существенное расхождение. "
        "🟨 **жёлтый** — конфликт назначения помещения. "
        "🟧 **оранжевый** — нет метки на чертеже (не с чем сравнить). "
        "⬜ **серый** — итоговая строка."
    )

    with st.expander("Почему так бывает в реальных проектах"):
        st.markdown("""
- **Разные определения площади.** БТИ считает «общую площадь» без балконов
  и с понижающим коэффициентом для лоджий (0,5). Дизайнер — по внутренним стенам.
- **Разные версии проекта.** Чертёж может быть до перепланировки, а экспликация — после.
- **ГКЛ-конструкции.** Гипсокартонные короба, ниши, встроенные шкафы по-разному
  учитываются в площади.
- **Кружки и таблица из разных источников.** Кружки часто ставятся для
  визуальной ориентировки, экспликация — официальный документ БТИ.

**В расчёте используем экспликацию** — по ней работает смета.
Чертёж — для визуализации геометрии.
        """)

with tab3:
    st.subheader("Открытые вопросы к заказчику")
    st.markdown(
        "Позиции, по которым система **не может принять решение "
        "автоматически**. Требуется уточнение от заказчика."
    )

    questions = get_open_questions()
    if not questions:
        st.success("Открытых вопросов нет — данные согласованы.")
    else:
        for i, q in enumerate(questions, 1):
            with st.expander(f"{i}. [{q['type']}] {q.get('помещение', '')}", expanded=True):
                st.markdown(f"**Вопрос:** {q['вопрос']}")
                if q.get("возможные_причины"):
                    st.markdown("**Возможные причины:**")
                    for cause in q["возможные_причины"]:
                        st.markdown(f"- {cause}")
                if q.get("действие"):
                    st.info(f"→ {q['действие']}")

with tab4:
    st.subheader(f"Расчёт объёмов (источник: {source})")
    st.dataframe(DF_VOL, use_container_width=True, hide_index=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Помещений", len(ROOMS))
    c2.metric("Площадь пола", f"{DF_VOL.iloc[:-1]['Площадь пола (м²)'].sum():.2f} м²")
    c3.metric("Площадь стен", f"{DF_VOL.iloc[:-1]['Площадь стен (м²)'].sum():.2f} м²")
    c4.metric("Периметр", f"{DF_VOL.iloc[:-1]['Периметр (м)'].sum():.2f} м")

    c5, c6 = st.columns(2)
    c5.metric("Двери", int(DF_VOL.iloc[:-1]["Двери (шт)"].sum()))
    c6.metric("Окна", int(DF_VOL.iloc[:-1]["Окна (шт)"].sum()))

with tab5:
    st.subheader("Excel-смета")
    st.markdown(
        "Файл содержит **три листа**:\n"
        "- «Объёмы работ» — основные расчёты\n"
        "- «Сверка источников» — сравнение экспликации и чертежа с подсветкой\n"
        "- «Открытые вопросы» — список вопросов к заказчику"
    )
    excel_buf = export_excel_bytes(DF_VOL, df_cmp, source=source)
    st.download_button(
        "⬇️ Скачать volumes.xlsx",
        data=excel_buf,
        file_name="volumes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab6:
    st.subheader("О решении")
    st.markdown("""
**Что автоматизировано:**
- Построение 3D-модели: пол, стены, потолок (опционально), мебель.
- Расстановка мебели по правилам для каждого типа помещения.
- Дедупликация стен: общая стена между комнатами рисуется один раз.
- Расчёт объёмов: площадь пола, стен (за вычетом проёмов), потолка, периметр.
- **Сверка источников:** автоматическое сравнение экспликации и чертежа.
- **Выявление конфликтов назначения** помещений.
- Выгрузка в Excel с тремя листами.

**Что сделано вручную:**
- Первичная оцифровка координат помещений с PDF (точность ±10–15 %).
- Выбор источника площадей (по умолчанию — экспликация).

**Обнаруженные проблемы в исходных данных:**
- Расхождения площадей: Кухня-Гостиная, Спальня, Спальня №2, Прихожая.
- Конфликт назначения: помещение «Спальня» на плане (кружок ⑥) выглядит
  как балкон по форме и толщине стен, но в экспликации записано как спальня.
- «Балкон» (9,50 м²) не имеет собственного кружка на плане.

**Стек:** Python 3.13, uv, Streamlit, NumPy, Pandas, Plotly, trimesh, openpyxl.

**Развитие:**
1. LLM Vision API (GPT-4o / Claude) — авто-извлечение координат стен и помещений.
2. Computer Vision (OpenCV / YOLO) — детекция стен и проёмов.
3. Автоматическая сверка экспликации, размерных линий и площадей.
4. Вырезание дверных и оконных проёмов через `trimesh.boolean.difference`.
5. Экспорт в IFC через `ifcopenshell` для BIM-совместимости.
6. Связка с расценками заказчика → автоматическая смета.
""")

with st.sidebar:
    st.divider()
    st.header("Параметры")
    st.write(f"**Высота потолков:** {WALL_HEIGHT} м")
    st.write(f"**Помещений:** {len(ROOMS)}")
    st.write(f"**Мебель:** {sum(len(r['furn']) for r in ROOMS.values())} объектов")
    st.divider()
    st.caption("Стены 80 мм, проёмы стандартные.")