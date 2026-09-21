# core.py
# Общая логика: данные планировки, сверка источников, геометрия (полигоны),
# расстановка мебели, сборка 3D-модели, расчёт объёмов, экспорт в .obj и .xlsx.

import io
import numpy as np
import pandas as pd
import trimesh
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill


# ============================================================
# 1. ДАННЫЕ ПЛАНИРОВКИ
# ============================================================
WALL_HEIGHT = 2.9
WALL_THICKNESS_INNER = 0.08
WALL_THICKNESS_OUTER = 0.25

EXPLICATION = {
    "Прихожая":       9.87,
    "Санузел":        2.09,
    "Ванная":         6.73,
    "Кухня-Гостиная": 35.45,
    "Спальня":        12.42,
    "Спальня №2":     17.50,
    "Балкон":         9.50,
}

DRAWING_MARKS = {
    "Прихожая":       3.40,     # ①
    "Санузел":        2.09,     # ④
    "Ванная":         6.80,     # ⑤
    "Кухня-Гостиная": 52.07,    # ②
    "Спальня":        20.72,    # ③
    "Спальня №2":     None,
    "Балкон":         9.74,     # ⑥
}

NAMING_CONFLICTS = {}

# --- Геометрия помещений ---
# Кухня-Гостиная: квадрат 5.93 × 5.93 с вырезом под Прихожую и Санузел.
# Спальня: L-форма 3.50 × 5.93 с вырезом под Ванную в правом нижнем углу.
# Балкон: снаружи сверху, с ограждением 1.1 м (wall_h).
ROOMS = {
    "Кухня-Гостиная": {
        "poly": [
            (0.00, 0.00),
            (3.25, 0.00),
            (3.25, 2.10),
            (5.93, 2.10),
            (5.93, 5.93),
            (0.00, 5.93),
        ],
        "h": 2.894, "doors": 2, "windows": 2,
        "furn": ["kitchen", "sofa", "table", "chair", "chair"],
        "floor_type": "паркет",
    },
    "Прихожая": {
        # Мебель убрана — только проход.
        "poly": [
            (4.30, 0.00),
            (5.93, 0.00),
            (5.93, 2.10),
            (4.30, 2.10),
        ],
        "h": 2.471, "doors": 2, "windows": 0,
        "furn": [],
        "floor_type": "плитка",
    },
    "Санузел": {
        "poly": [
            (3.25, 0.00),
            (4.30, 0.00),
            (4.30, 2.10),
            (3.25, 2.10),
        ],
        "h": 2.376, "doors": 1, "windows": 0,
        "furn": ["toilet"],
        "floor_type": "плитка",
    },
    "Балкон": {
        "poly": [
            (1.20, 5.93),
            (5.20, 5.93),
            (5.20, 7.53),
            (1.20, 7.53),
        ],
        "h": 2.90,
        "wall_h": 1.10,
        "doors": 1, "windows": 3,
        "furn": [],
        "floor_type": "плитка",
    },
    "Спальня": {
        "poly": [
            (5.93, 0.00),
            (7.23, 0.00),
            (7.23, 2.90),
            (9.43, 2.90),
            (9.43, 5.93),
            (5.93, 5.93),
        ],
        "h": 2.891, "doors": 1, "windows": 2,
        "furn": ["bed", "wardrobe"],
        "floor_type": "паркет",
    },
    "Ванная": {
        "poly": [
            (7.23, 0.00),
            (9.43, 0.00),
            (9.43, 2.90),
            (7.23, 2.90),
        ],
        "h": 2.900, "doors": 1, "windows": 0,
        "furn": ["tub", "toilet"],
        "floor_type": "плитка",
    },
}

FURNITURE = {
    "bed":      (1.6, 2.0, 0.5),
    "sofa":     (2.2, 0.9, 0.8),
    "wardrobe": (1.5, 0.6, 2.2),
    "table":    (1.2, 0.8, 0.75),
    "chair":    (0.5, 0.5, 0.9),
    "kitchen":  (3.0, 0.6, 0.9),
    "tub":      (1.7, 0.8, 0.6),
    "toilet":   (0.4, 0.6, 0.8),
}

# Человеко-читаемые названия мебели — для тултипов
FURNITURE_RU = {
    "bed":      "Кровать",
    "sofa":     "Диван",
    "wardrobe": "Шкаф",
    "table":    "Стол",
    "chair":    "Стул",
    "kitchen":  "Кухня",
    "tub":      "Ванна",
    "toilet":   "Унитаз",
}

COLOR_MAP = {
    "floor_wood": "#c8a97e",
    "floor_tile": "#a8b8c8",
    "ceiling":    "#f5f5f5",
    "wall":       "#b8b8b8",
    "furniture":  "#4a7c59",
}


# ============================================================
# 2. ГЕОМЕТРИЯ (полигоны)
# ============================================================
def poly_signed_area(poly):
    """Знаковая площадь полигона. Положительная — CCW, отрицательная — CW."""
    x = np.array([p[0] for p in poly])
    y = np.array([p[1] for p in poly])
    return 0.5 * (np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1)))


def poly_area(poly):
    """Модуль площади полигона."""
    return abs(poly_signed_area(poly))


def poly_perimeter(poly):
    """Периметр полигона."""
    p = np.array(poly + [poly[0]])
    return float(np.sum(np.linalg.norm(np.diff(p, axis=0), axis=1)))


def poly_centroid(poly):
    """
    Центроид полигона.
    Используется ЗНАКОВАЯ площадь, чтобы центроид был корректным
    и для CW-, и для CCW-полигонов.
    """
    x = np.array([p[0] for p in poly])
    y = np.array([p[1] for p in poly])
    A = poly_signed_area(poly)
    if abs(A) < 1e-9:
        return float(np.mean(x)), float(np.mean(y))
    cx = np.sum((x + np.roll(x, -1)) * (x * np.roll(y, -1) - np.roll(x, -1) * y)) / (6 * A)
    cy = np.sum((y + np.roll(y, -1)) * (x * np.roll(y, -1) - np.roll(x, -1) * y)) / (6 * A)
    return float(cx), float(cy)


def polygon_to_faces(poly, z=0.0, flip=False):
    """
    Триангуляция полигона веером от центроида.
    Корректно для CCW-полигонов, у которых центроид видит все вершины.
    """
    cx, cy = poly_centroid(poly)
    n = len(poly)
    verts = np.array([[p[0], p[1], z] for p in poly] + [[cx, cy, z]])
    faces = np.array([[i, (i + 1) % n, n] for i in range(n)])
    if flip:
        faces = faces[:, ::-1]
    return verts, faces


def box_mesh(cx, cy, w, d, h, z0=0.0):
    """Прямоугольный параллелепипед."""
    x0, x1 = cx - w / 2, cx + w / 2
    y0, y1 = cy - d / 2, cy + d / 2
    v = np.array([
        [x0, y0, z0], [x1, y0, z0], [x1, y1, z0], [x0, y1, z0],
        [x0, y0, z0 + h], [x1, y0, z0 + h], [x1, y1, z0 + h], [x0, y1, z0 + h],
    ])
    f = np.array([
        [0, 1, 2], [0, 2, 3], [4, 6, 5], [4, 7, 6],
        [0, 1, 5], [0, 5, 4], [1, 2, 6], [1, 6, 5],
        [2, 3, 7], [2, 7, 6], [3, 0, 4], [3, 4, 7],
    ])
    return v, f


def wall_mesh(p0, p1, height, thickness):
    """Стена между двумя точками."""
    x0, y0 = p0
    x1, y1 = p1
    dx, dy = x1 - x0, y1 - y0
    length = np.hypot(dx, dy)
    if length < 1e-6:
        return None
    nx, ny = -dy / length * thickness / 2, dx / length * thickness / 2
    v = np.array([
        [x0 + nx, y0 + ny, 0], [x0 - nx, y0 - ny, 0],
        [x1 - nx, y1 - ny, 0], [x1 + nx, y1 + ny, 0],
        [x0 + nx, y0 + ny, height], [x0 - nx, y0 - ny, height],
        [x1 - nx, y1 - ny, height], [x1 + nx, y1 + ny, height],
    ])
    f = np.array([
        [0, 2, 1], [0, 3, 2], [4, 5, 6], [4, 6, 7],
        [0, 1, 5], [0, 5, 4], [1, 2, 6], [1, 6, 5],
        [2, 3, 7], [2, 7, 6], [3, 0, 4], [3, 4, 7],
    ])
    return v, f


# ============================================================
# 3. ОПРЕДЕЛЕНИЕ ТИПА РЕБРА
# ============================================================
def get_apartment_bbox(rooms):
    """Bounding box всей квартиры по всем полигонам."""
    all_x, all_y = [], []
    for data in rooms.values():
        for x, y in data["poly"]:
            all_x.append(x)
            all_y.append(y)
    return min(all_x), min(all_y), max(all_x), max(all_y)


def is_outer_edge(p1, p2, bbox, tol=0.05):
    """Ребро считается наружным, если его середина лежит на границе bbox."""
    mx = (p1[0] + p2[0]) / 2
    my = (p1[1] + p2[1]) / 2
    x_min, y_min, x_max, y_max = bbox
    return (
        abs(mx - x_min) < tol or
        abs(mx - x_max) < tol or
        abs(my - y_min) < tol or
        abs(my - y_max) < tol
    )


# ============================================================
# 4. РАССТАНОВКА МЕБЕЛИ
# ============================================================
def place_furniture(name, poly, items):
    """
    Расстановка мебели с учётом формы помещения и проходов.
    Главное правило: не перекрывать проходы между помещениями.
    """
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    placed = []
    chair_idx = 0

    for item in items:
        fw, fd, fh = FURNITURE[item]

        if name == "Кухня-Гостиная":
            # Кухня — вдоль северной стены, левая часть
            if item == "kitchen":
                placed.append((item, 1.80, y1 - fd / 2 - 0.1))     # x 0.30–3.30
            # Диван — в левой зоне, спинка к югу
            elif item == "sofa":
                placed.append((item, 1.60, 3.75))                  # x 0.50–2.70
            # Стол — правее дивана, но не в проходе
            elif item == "table":
                placed.append((item, 4.00, 4.40))                  # x 3.40–4.60
            # Стулья — по бокам стола
            elif item == "chair":
                positions = [(3.05, 4.40), (4.95, 4.40)]
                px, py = positions[chair_idx % len(positions)]
                chair_idx += 1
                placed.append((item, px, py))

        elif name == "Спальня":
            # Кровать — изголовьем к северной стене, сдвинута к востоку
            if item == "bed":
                placed.append((item, 7.80, y1 - fd / 2 - 0.1))     # x 7.00–8.60
            # Шкаф — у восточной стены, ниже кровати
            elif item == "wardrobe":
                placed.append((item, 8.20, 3.30))                  # x 8.83–9.43

        elif name == "Ванная":
            if item == "tub":
                placed.append((item, x0 + fw / 2 + 0.1, y1 - fd / 2 - 0.1))
            elif item == "toilet":
                placed.append((item, x1 - fw / 2 - 0.3, y0 + fd / 2 + 0.3))

        elif name == "Санузел":
            if item == "toilet":
                placed.append((item, x0 + fw / 2 + 0.2, y1 - fd / 2 - 0.2))

        # Прихожая — без мебели (только проход)

    return placed


# ============================================================
# 5. СБОРКА 3D-МОДЕЛИ
# ============================================================
def build_all(include_ceiling=False):
    """
    Собирает 3D-модель.
    Возвращает (verts, faces, colors) — colors это массив имён цветов.
    Используется в export_obj_file.
    """
    all_v, all_f, all_c = [], [], []
    offset = 0

    def add(v, f, color):
        nonlocal offset
        all_v.append(v)
        all_f.append(f + offset)
        all_c.extend([color] * len(v))
        offset += len(v)

    bbox = get_apartment_bbox(ROOMS)

    # Полы
    for name, data in ROOMS.items():
        color = "floor_wood" if data.get("floor_type") == "паркет" else "floor_tile"
        v, f = polygon_to_faces(data["poly"], z=0.01, flip=False)
        add(v, f, color)

    # Потолки (опционально)
    if include_ceiling:
        for name, data in ROOMS.items():
            v, f = polygon_to_faces(data["poly"], z=data["h"], flip=True)
            add(v, f, "ceiling")

    # Стены
    edge_seen = set()
    for name, data in ROOMS.items():
        poly = data["poly"]
        h = data.get("wall_h", data["h"])
        n = len(poly)
        for i in range(n):
            p1 = poly[i]
            p2 = poly[(i + 1) % n]
            key = (round(min(p1[0], p2[0]), 3),
                   round(min(p1[1], p2[1]), 3),
                   round(max(p1[0], p2[0]), 3),
                   round(max(p1[1], p2[1]), 3))
            if key in edge_seen:
                continue
            edge_seen.add(key)
            thickness = (WALL_THICKNESS_OUTER
                         if is_outer_edge(p1, p2, bbox)
                         else WALL_THICKNESS_INNER)
            res = wall_mesh(p1, p2, h, thickness=thickness)
            if res:
                add(*res, "wall")

    # Мебель
    for name, data in ROOMS.items():
        for item, x, y in place_furniture(name, data["poly"], data["furn"]):
            fw, fd, fh = FURNITURE[item]
            v, f = box_mesh(x, y, fw, fd, fh, z0=0.02)
            add(v, f, "furniture")

    return np.vstack(all_v), np.vstack(all_f), np.array(all_c)


# ============================================================
# 5b. СБОРКА 3D-МОДЕЛИ С МЕТАДАННЫМИ (для тултипов)
# ============================================================
def build_all_with_meta(include_ceiling=False):
    """
    Возвращает (verts, faces, meta).
    meta — список словарей, по одному на каждую вершину:
      {"color": "wall", "room": "Кухня-Гостиная", "category": "wall",
       "label": "Кухня-Гостиная — Стены"}
    Используется в Streamlit для подписей при наведении.
    """
    all_v, all_f, meta = [], [], []
    offset = 0

    def add(v, f, color, room, category, label):
        nonlocal offset
        all_v.append(v)
        all_f.append(f + offset)
        meta.extend([{
            "color": color,
            "room": room,
            "category": category,
            "label": label,
        }] * len(v))
        offset += len(v)

    bbox = get_apartment_bbox(ROOMS)

    # Полы
    for name, data in ROOMS.items():
        color = "floor_wood" if data.get("floor_type") == "паркет" else "floor_tile"
        v, f = polygon_to_faces(data["poly"], z=0.01, flip=False)
        add(v, f, color, name, "floor", f"{name} — Пол")

    # Потолки (опционально)
    if include_ceiling:
        for name, data in ROOMS.items():
            v, f = polygon_to_faces(data["poly"], z=data["h"], flip=True)
            add(v, f, "ceiling", name, "ceiling", f"{name} — Потолок")

    # Стены
    edge_seen = set()
    for name, data in ROOMS.items():
        poly = data["poly"]
        h = data.get("wall_h", data["h"])
        n = len(poly)
        for i in range(n):
            p1 = poly[i]
            p2 = poly[(i + 1) % n]
            key = (round(min(p1[0], p2[0]), 3),
                   round(min(p1[1], p2[1]), 3),
                   round(max(p1[0], p2[0]), 3),
                   round(max(p1[1], p2[1]), 3))
            if key in edge_seen:
                continue
            edge_seen.add(key)
            thickness = (WALL_THICKNESS_OUTER
                         if is_outer_edge(p1, p2, bbox)
                         else WALL_THICKNESS_INNER)
            res = wall_mesh(p1, p2, h, thickness=thickness)
            if res:
                wall_type = "Наружная стена" if thickness == WALL_THICKNESS_OUTER else "Стена"
                add(*res, "wall", name, "wall", f"{name} — {wall_type}")

    # Мебель
    for name, data in ROOMS.items():
        for item, x, y in place_furniture(name, data["poly"], data["furn"]):
            fw, fd, fh = FURNITURE[item]
            v, f = box_mesh(x, y, fw, fd, fh, z0=0.02)
            item_ru = FURNITURE_RU.get(item, item)
            label = f"{name} — {item_ru} ({fw:.1f}×{fd:.1f} м)"
            add(v, f, "furniture", name, "furniture", label)

    return np.vstack(all_v), np.vstack(all_f), meta


# ============================================================
# 6. СВЕРКА ИСТОЧНИКОВ
# ============================================================
def compare_sources(tolerance_pct: float = 5.0) -> pd.DataFrame:
    rows = []
    for name in EXPLICATION:
        exp = EXPLICATION.get(name)
        draw = DRAWING_MARKS.get(name)

        if exp is not None and draw is not None:
            diff = draw - exp
            diff_pct = (diff / exp) * 100 if exp else 0
            base_status = "OK" if abs(diff_pct) <= tolerance_pct else "РАСХОЖДЕНИЕ"
        elif exp is not None and draw is None:
            diff, diff_pct, base_status = None, None, "нет на чертеже"
        else:
            diff, diff_pct, base_status = None, None, "—"

        status = base_status
        if name in NAMING_CONFLICTS:
            status = f"{base_status} ⚠ конфликт назначения"

        rows.append({
            "Помещение": name,
            "Экспликация (м²)": exp,
            "Чертёж (м²)": draw,
            "Разница (м²)": round(diff, 2) if diff is not None else None,
            "Разница (%)": round(diff_pct, 1) if diff_pct is not None else None,
            "Статус": status,
        })

    exp_sum = sum(v for v in EXPLICATION.values() if v is not None)
    draw_sum = sum(v for v in DRAWING_MARKS.values() if v is not None)
    rows.append({
        "Помещение": "ИТОГО",
        "Экспликация (м²)": round(exp_sum, 2),
        "Чертёж (м²)": round(draw_sum, 2),
        "Разница (м²)": round(draw_sum - exp_sum, 2),
        "Разница (%)": round((draw_sum - exp_sum) / exp_sum * 100, 1),
        "Статус": "—",
    })
    return pd.DataFrame(rows)


def get_open_questions() -> list:
    questions = []
    df = compare_sources()

    for _, row in df.iterrows():
        if row["Помещение"] == "ИТОГО":
            continue
        if isinstance(row["Статус"], str) and "РАСХОЖДЕНИЕ" in row["Статус"]:
            if "конфликт назначения" not in row["Статус"]:
                questions.append({
                    "type": "Расхождение площадей",
                    "помещение": row["Помещение"],
                    "вопрос": (
                        f"Экспликация: {row['Экспликация (м²)']} м², "
                        f"чертёж: {row['Чертёж (м²)']} м² "
                        f"(разница {row['Разница (%)']} %). "
                        f"Какую цифру брать в смету?"
                    ),
                })

    for name, info in NAMING_CONFLICTS.items():
        questions.append({
            "type": "Конфликт назначения",
            "помещение": name,
            "вопрос": info["issue"],
            "возможные_причины": info["possible_causes"],
            "действие": info["action"],
        })

    for _, row in df.iterrows():
        if row["Помещение"] == "ИТОГО":
            continue
        if row["Статус"] == "нет на чертеже":
            questions.append({
                "type": "Нет метки на чертеже",
                "помещение": row["Помещение"],
                "вопрос": (
                    f"В экспликации «{row['Помещение']}» = {row['Экспликация (м²)']} м², "
                    f"но на плане нет кружка и нет отдельного помещения "
                    f"с такой площадью. Это опечатка или помещение "
                    f"пропущено на чертеже?"
                ),
            })

    return questions


def get_selected_areas(source: str = "explication") -> dict:
    if source == "explication":
        return {k: v for k, v in EXPLICATION.items() if v is not None}

    if source == "drawing":
        result = {}
        for name in EXPLICATION:
            draw = DRAWING_MARKS.get(name)
            result[name] = draw if draw is not None else EXPLICATION[name]
        return result

    if source == "geometry":
        return {name: poly_area(d["poly"]) for name, d in ROOMS.items()}

    raise ValueError(f"Unknown source: {source}")


# ============================================================
# 7. РАСЧЁТ ОБЪЁМОВ
# ============================================================
def compute_volumes(source: str = "explication") -> pd.DataFrame:
    areas = get_selected_areas(source)
    rows = []
    for name, d in ROOMS.items():
        area = areas.get(name, poly_area(d["poly"]))
        perim = poly_perimeter(d["poly"])
        h = d["h"]
        doors = d.get("doors", 0)
        windows = d.get("windows", 0)
        openings = doors * 0.9 * 2.1 + windows * 1.5 * 1.5
        wall_area = perim * h - openings

        rows.append({
            "Помещение": name,
            "Площадь пола (м²)": round(area, 2),
            "Периметр (м)": round(perim, 2),
            "Высота (м)": h,
            "Площадь стен (м²)": round(wall_area, 2),
            "Площадь потолка (м²)": round(area, 2),
            "Двери (шт)": doors,
            "Окна (шт)": windows,
            "Тип пола": d.get("floor_type", "—"),
        })

    rows.append({
        "Помещение": "ИТОГО",
        "Площадь пола (м²)": round(sum(r["Площадь пола (м²)"] for r in rows), 2),
        "Периметр (м)": round(sum(r["Периметр (м)"] for r in rows), 2),
        "Высота (м)": None,
        "Площадь стен (м²)": round(sum(r["Площадь стен (м²)"] for r in rows), 2),
        "Площадь потолка (м²)": round(sum(r["Площадь потолка (м²)"] for r in rows), 2),
        "Двери (шт)": int(sum(r["Двери (шт)"] for r in rows)),
        "Окна (шт)": int(sum(r["Окна (шт)"] for r in rows)),
        "Тип пола": "—",
    })
    return pd.DataFrame(rows)


# ============================================================
# 8. ЭКСПОРТ В EXCEL
# ============================================================
def export_excel_bytes(df_volumes=None, df_compare=None,
                       source: str = "explication") -> io.BytesIO:
    df_volumes = df_volumes if df_volumes is not None else compute_volumes(source)
    df_compare = df_compare if df_compare is not None else compare_sources()
    questions = get_open_questions()

    buf = io.BytesIO()
    _write_workbook(df_volumes, df_compare, questions, buf, source)
    buf.seek(0)
    return buf


def export_excel_file(path="volumes.xlsx", df_volumes=None, df_compare=None,
                      source: str = "explication"):
    df_volumes = df_volumes if df_volumes is not None else compute_volumes(source)
    df_compare = df_compare if df_compare is not None else compare_sources()
    questions = get_open_questions()
    _write_workbook(df_volumes, df_compare, questions, path, source)
    print(f"✓ Объёмы сохранены: {path}")
    return df_volumes


def _write_workbook(df_volumes, df_compare, questions, target, source):
    wb = Workbook()

    # --- Лист 1: Объёмы работ ---
    ws = wb.active
    ws.title = "Объёмы работ"
    ws.append(["Помещение", "Показатель", "Ед. изм.", "Количество"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for _, row in df_volumes.iterrows():
        n = row["Помещение"]
        is_total = (n == "ИТОГО")
        fill = PatternFill("solid", fgColor="DDDDDD") if is_total else None

        lines = [
            [n, "Площадь пола",    "м²", row["Площадь пола (м²)"]],
            [n, "Площадь стен",    "м²", row["Площадь стен (м²)"]],
            [n, "Площадь потолка", "м²", row["Площадь потолка (м²)"]],
            [n, "Периметр",        "м",  row["Периметр (м)"]],
            [n, "Двери",           "шт", row["Двери (шт)"]],
            [n, "Окна",            "шт", row["Окна (шт)"]],
        ]
        for line in lines:
            ws.append(line)
            if fill:
                for cell in ws[ws.max_row]:
                    cell.fill = fill
                    cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 15

    ws.append([])
    source_label = {
        "explication": "Экспликация (официальный документ)",
        "drawing":     "Чертёж (кружки на плане)",
        "geometry":    "Геометрия 3D-модели (полигоны)",
    }[source]
    ws.append(["Источник площадей:", source_label])

    # --- Лист 2: Сверка источников ---
    ws2 = wb.create_sheet("Сверка источников")
    ws2.append(list(df_compare.columns))
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    fill_warn = PatternFill("solid", fgColor="FFC7CE")
    fill_ok = PatternFill("solid", fgColor="C6EFCE")
    fill_conflict = PatternFill("solid", fgColor="FFEB9C")
    fill_missing = PatternFill("solid", fgColor="FFD8A8")
    fill_total = PatternFill("solid", fgColor="DDDDDD")

    for _, row in df_compare.iterrows():
        ws2.append(list(row.values))
        status = str(row["Статус"])
        r = ws2.max_row
        if row["Помещение"] == "ИТОГО":
            for cell in ws2[r]:
                cell.fill = fill_total
                cell.font = Font(bold=True)
        elif "конфликт назначения" in status:
            for cell in ws2[r]:
                cell.fill = fill_conflict
        elif "РАСХОЖДЕНИЕ" in status:
            for cell in ws2[r]:
                cell.fill = fill_warn
        elif status == "OK":
            for cell in ws2[r]:
                cell.fill = fill_ok
        elif "нет на чертеже" in status:
            for cell in ws2[r]:
                cell.fill = fill_missing

    for col, width in zip("ABCDEF", [22, 20, 18, 15, 15, 30]):
        ws2.column_dimensions[col].width = width

    # --- Лист 3: Открытые вопросы ---
    ws3 = wb.create_sheet("Открытые вопросы")
    ws3.append(["Тип", "Помещение", "Вопрос", "Возможные причины", "Действие"])
    for cell in ws3[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for q in questions:
        ws3.append([
            q.get("type", ""),
            q.get("помещение", ""),
            q.get("вопрос", ""),
            "; ".join(q.get("возможные_причины", [])) if q.get("возможные_причины") else "",
            q.get("действие", ""),
        ])

    for col, width in zip("ABCDE", [25, 20, 60, 60, 30]):
        ws3.column_dimensions[col].width = width

    if isinstance(target, io.BytesIO):
        wb.save(target)
    else:
        wb.save(target)


# ============================================================
# 9. ЭКСПОРТ 3D-МОДЕЛИ В .OBJ
# ============================================================
def export_obj_file(path="apartment.obj", include_ceiling=False):
    """
    Сохраняет 3D-модель как .obj — универсальный формат.
    Открывается в Blender, 3ds Max, Maya, SketchUp, онлайн-вьюерах.
    """
    verts, faces, _ = build_all(include_ceiling=include_ceiling)
    mesh = trimesh.Trimesh(vertices=verts, faces=faces, process=False)
    mesh.export(path)
    print(f"✓ 3D-модель сохранена: {path}")
    return mesh