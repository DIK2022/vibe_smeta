# build_model.py
# CLI: создаёт apartment.obj и volumes.xlsx.
# Запуск: uv run build_model.py

from core import (
    export_obj_file, export_excel_file,
    compare_sources, compute_volumes, get_open_questions,
)


def main():
    print("=" * 78)
    print("СВЕРКА ИСХОДНЫХ ДАННЫХ: ЭКСПЛИКАЦИЯ vs ЧЕРТЁЖ")
    print("=" * 78)
    df_cmp = compare_sources()
    print(df_cmp.to_string(index=False))
    print()

    n_warn = df_cmp["Статус"].astype(str).str.contains("РАСХОЖДЕНИЕ").sum()
    if n_warn:
        print(f"⚠  Найдено расхождений: {n_warn}")
        print("   В расчёте используется ЭКСПЛИКАЦИЯ как официальный источник.")
    print()

    # Открытые вопросы
    questions = get_open_questions()
    if questions:
        print("=" * 78)
        print("ОТКРЫТЫЕ ВОПРОСЫ К ЗАКАЗЧИКУ")
        print("=" * 78)
        for i, q in enumerate(questions, 1):
            print(f"\n{i}. [{q['type']}] {q.get('помещение', '')}")
            print(f"   {q['вопрос']}")
            if q.get("возможные_причины"):
                print("   Возможные причины:")
                for cause in q["возможные_причины"]:
                    print(f"     - {cause}")
            if q.get("действие"):
                print(f"   → {q['действие']}")
        print()

    print("=" * 78)
    print("Строю 3D-модель квартиры...")
    print("=" * 78)
    export_obj_file("apartment.obj")

    print()
    print("=" * 78)
    print("Считаю объёмы работ (источник площадей: экспликация)...")
    print("=" * 78)
    df = export_excel_file("volumes.xlsx", source="explication")

    print()
    print("Итоговая таблица объёмов:")
    print(df.to_string(index=False))
    print()
    print("Готово. Файлы: apartment.obj, volumes.xlsx")


if __name__ == "__main__":
    main()