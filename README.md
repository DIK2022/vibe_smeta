# PDF-планировка → 3D-модель → Excel-смета

Прототип системы, которая превращает планировку квартиры из PDF
в упрощённую 3D-модель, расставляет мебель, считает объёмы работ
и выгружает результат в Excel — со **сверкой источников данных**.

## Стек

Python 3.13, uv, NumPy, Pandas, trimesh, openpyxl.
Опционально: Streamlit + Plotly для интерактивного просмотра.

## Запуск

```bash
uv python install 3.13
uv sync

# Основной результат по ТЗ: apartment.glb + volumes.xlsx
uv run build_model.py

# Интерактивная демонстрация (бонус, не требуется ТЗ)
uv run streamlit run app.py# vibe_smeta
