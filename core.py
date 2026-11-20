"""Данные, разделы карточки станции, расчёты и экспорт (Excel / Word)."""
from __future__ import annotations

import datetime as dt
import io
import json
import math
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

DATA = Path(__file__).parent / "data"
MONTHS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
CUR = "СП 131.13330.2025"
PREV_EDITIONS = ["СП 131.13330.2020", "СП 131.13330.2018", "СНиП 23-01-99*"]

# (ключ, № графы в таблице СП, наименование, ед.)
COLD = [
    ("c_day098", 2, "Температура воздуха наиболее холодных суток, обеспеченностью 0,98", "°С"),
    ("c_day092", 3, "Температура воздуха наиболее холодных суток, обеспеченностью 0,92", "°С"),
    ("c_5d098", 4, "Температура воздуха наиболее холодной пятидневки, обеспеченностью 0,98", "°С"),
    ("c_5d092", 5, "Температура воздуха наиболее холодной пятидневки, обеспеченностью 0,92", "°С"),
    ("c_t094", 6, "Температура воздуха обеспеченностью 0,94", "°С"),
    ("c_absmin", 7, "Абсолютная минимальная температура воздуха", "°С"),
    ("c_ampl", 8, "Средняя суточная амплитуда температуры воздуха наиболее холодного месяца", "°С"),
    ("c_n0", 9, "Продолжительность периода со средней суточной температурой воздуха ≤ 0 °С", "сут"),
    ("c_t0", 10, "Средняя температура воздуха периода со средней суточной температурой ≤ 0 °С", "°С"),
    ("c_n8", 11, "Продолжительность периода со средней суточной температурой воздуха ≤ 8 °С", "сут"),
    ("c_t8", 12, "Средняя температура воздуха периода со средней суточной температурой ≤ 8 °С", "°С"),
    ("c_n10", 13, "Продолжительность периода со средней суточной температурой воздуха ≤ 10 °С", "сут"),
    ("c_t10", 14, "Средняя температура воздуха периода со средней суточной температурой ≤ 10 °С", "°С"),
    ("c_rh", 15, "Средняя месячная относительная влажность воздуха наиболее холодного месяца", "%"),
    ("c_rh15", 16, "Средняя месячная относительная влажность воздуха в 15 ч наиболее холодного месяца", "%"),
    ("c_prec", 17, "Количество осадков за ноябрь–март", "мм"),
    ("c_wdir", 18, "Преобладающее направление ветра за декабрь–февраль", "—"),
    ("c_wmax", 19, "Максимальная из средних скоростей ветра по румбам за январь", "м/с"),
    ("c_wavg", 20, "Средняя скорость ветра за период со средней суточной температурой воздуха ≤ 8 °С", "м/с"),
]
WARM = [
    ("w_p", 2, "Барометрическое давление", "гПа"),
    ("w_t095", 3, "Температура воздуха обеспеченностью 0,95", "°С"),
    ("w_t098", 4, "Температура воздуха обеспеченностью 0,98", "°С"),
    ("w_tmax_avg", 5, "Средняя максимальная температура воздуха наиболее теплого месяца", "°С"),
    ("w_absmax", 6, "Абсолютная максимальная температура воздуха", "°С"),
    ("w_ampl", 7, "Средняя суточная амплитуда температуры воздуха наиболее теплого месяца", "°С"),
    ("w_rh", 8, "Средняя месячная относительная влажность воздуха наиболее теплого месяца", "%"),
    ("w_rh15", 9, "Средняя месячная относительная влажность воздуха в 15 ч наиболее теплого месяца", "%"),
    ("w_prec", 10, "Количество осадков за апрель–октябрь", "мм"),
    ("w_precmax", 11, "Суточный максимум осадков", "мм"),
    ("w_wdir", 12, "Преобладающее направление ветра за июнь–август", "—"),
    ("w_wmin", 13, "Минимальная из средних скоростей ветра по румбам за июль", "м/с"),
]

SOILS = {
    "суглинки и глины": 0.23,
    "супеси, пески мелкие и пылеватые": 0.28,
    "пески гравелистые, крупные и средней крупности": 0.30,
    "крупнообломочные грунты": 0.34,
}

# справочные таблицы 13.1, 13.2, А.1, Б.1 (СП 131) и табл. 5.2 СП 22 (kh)
REF: dict = json.loads((DATA / "ref_tables.json").read_text(encoding="utf-8")) \
    if (DATA / "ref_tables.json").exists() else {}
MONTH_NAMES = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август",
               "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

SOURCES = (
    "СП 131.13330.2025 «Строительная климатология» — основной источник. "
    "Справочно: СП 131.13330.2020, СП 131.13330.2018, СНиП 23-01-99*; "
    "СП 20.13330.2016, Прил. К (вес снегового покрова); "
    "карты ОСР-2016 по перечню Прил. А СП 14.13330.2018 (изм. 1)."
)


# ------------------------------------------------------------------ данные
def load_data():
    st = pd.read_csv(DATA / "stations.csv", encoding="utf-8")
    prev = pd.read_csv(DATA / "prev_editions.csv", encoding="utf-8")
    osr = pd.read_csv(DATA / "osr2016.csv", encoding="utf-8", dtype={"A": str, "B": str, "C": str,
                                                                     "A_2015": str, "B_2015": str, "C_2015": str})
    for c in [c for c in st.columns if c.startswith("osr_") and c not in ("osr_name", "osr_dist_km")]:
        st[c] = st[c].map(_ball)
    st["climate_sub"] = st.apply(lambda r: ", ".join(climate_subregions(rec(r))) or None, axis=1)
    st["mt"] = st.apply(lambda r: frost_index(rec(r)), axis=1)
    st["search"] = (st["name"].fillna("") + " " + st["address"].fillna("") + " " + st["region"].fillna("")
                    ).str.lower().str.replace("ё", "е")
    return st, prev, osr


def rec(row) -> dict:
    """Строка DataFrame -> dict, NaN -> None, numpy-скаляры -> python."""
    out = {}
    for k, v in dict(row).items():
        if v is None or (isinstance(v, (float, np.floating)) and math.isnan(v)):
            out[k] = None
        elif isinstance(v, np.integer):
            out[k] = int(v)
        elif isinstance(v, np.floating):
            out[k] = float(v)
        else:
            out[k] = v
    return out


def _ball(v):
    """Балл как текст: 9.0 -> '9', '<6' без изменений, пусто -> None."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        return str(int(v)) if float(v).is_integer() else str(v)
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def fmt(v):
    """Значение для отображения."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    if isinstance(v, (int, np.integer)):
        return str(v)
    if isinstance(v, (float, np.floating)):
        r = round(float(v), 2)
        return str(int(r)) if r.is_integer() else f"{r:g}"
    return str(v)


def haversine_km(lat1, lon1, lat2, lon2):
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dp, dl = p2 - p1, np.radians(np.asarray(lon2) - lon1)
    a = np.sin(dp / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dl / 2) ** 2
    return 2 * 6371.0 * np.arcsin(np.sqrt(a))


def parse_point(text: str):
    """'55.75, 37.61' / '55,75 37,61' -> (lat, lon) или None."""
    if not text:
        return None
    t = text.replace(";", " ").strip()
    parts = [p for p in t.replace(", ", " ").replace(",", ".").split() if p]
    if len(parts) != 2:
        return None
    try:
        lat, lon = float(parts[0]), float(parts[1])
    except ValueError:
        return None
    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return None
    return lat, lon


# ----------------------------------------------------------------- расчёты
def frost_index(r):
    """Мt — сумма абсолютных значений среднемесячных отрицательных температур (табл. 7.1)."""
    vals = [r.get(f"t_{m}") for m in MONTHS]
    if any(v is None for v in vals):
        return None
    return round(sum(-v for v in vals if v < 0), 1)


def winter_mean_negative(r):
    vals = [r.get(f"t_{m}") for m in MONTHS]
    neg = [v for v in vals if v is not None and v < 0]
    return round(sum(neg) / len(neg), 1) if neg else None


def climate_subregions(r):
    """Подрайоны, условиям которых удовлетворяют данные станции (табл. А.1 СП 131). Условно."""
    tj, tjul = r.get("t_I"), r.get("t_VII")
    vw = [r.get("v_XII"), r.get("v_I"), r.get("v_II")]
    rh = r.get("w_rh")
    rh15 = r.get("w_rh15")
    if tj is None or tjul is None:
        return []
    wind = sum(vw) / 3 if all(v is not None for v in vw) else None

    def rng(x, lo, hi):
        return x is not None and lo <= x <= hi

    windy = wind is not None and wind >= 5
    humid = rh is not None and rh > 75
    n0 = r.get("c_n0")
    long_cold = n0 is not None and n0 >= 190  # примечание к табл. А.1 для IД
    rules = [
        ("IА", tj <= -32 and rng(tjul, 4, 19)),
        ("IБ", tj <= -28 and windy and rng(tjul, 0, 13) and humid),
        ("IВ", rng(tj, -28, -14) and rng(tjul, 12, 21)),
        ("IГ", rng(tj, -28, -14) and windy and rng(tjul, 0, 14) and humid),
        ("IД", rng(tj, -32, -14) and rng(tjul, 10, 20) and long_cold),
        ("IIА", rng(tj, -14, -4) and windy and rng(tjul, 8, 12) and humid),
        ("IIБ", rng(tj, -5, -3) and windy and rng(tjul, 12, 21) and humid),
        ("IIВ", rng(tj, -14, -4) and rng(tjul, 12, 21)),
        ("IIГ", rng(tj, -14, -5) and windy and rng(tjul, 12, 21) and humid),
        ("IIIА", rng(tj, -20, -14) and rng(tjul, 21, 25)),
        ("IIIБ", rng(tj, -5, 2) and rng(tjul, 21, 25)),
        ("IIIВ", rng(tj, -14, -5) and rng(tjul, 21, 25)),
        ("IVА", rng(tj, -10, 2) and tjul >= 28),
        ("IVБ", rng(tj, 2, 6) and rng(tjul, 22, 28) and rh15 is not None and rh15 > 50),
        ("IVВ", rng(tj, 0, 2) and rng(tjul, 25, 28)),
        ("IVГ", rng(tj, -15, 0) and rng(tjul, 25, 28)),
    ]
    return [n for n, ok in rules if ok]


def severity_check(r):
    """Табл. Б.1: по каждому из 6 показателей — подзоны, в диапазон которых попадает станция."""
    vals = [r.get("c_absmin"), r.get("c_day098"), r.get("c_day092"), r.get("c_5d098"), r.get("c_5d092"),
            (r["c_n8"] * r["c_t8"]) if r.get("c_n8") is not None and r.get("c_t8") is not None else None]
    rows = REF["tB1"]["rows"]
    zones = {}
    for i in range(0, len(rows), 2):
        name, hi, lo = rows[i][0], rows[i][1:], rows[i + 1][1:]
        zones[name] = [v is not None and lo[k] <= v <= hi[k] for k, v in enumerate(vals)]
    return vals, zones


def _interp_lat(lats, values, lat):
    """Линейная интерполяция по широте; вне диапазона или при пропуске — None."""
    if lat is None or lat < lats[0] or lat > lats[-1]:
        return None
    for i in range(len(lats) - 1):
        if lats[i] <= lat <= lats[i + 1]:
            a, b = values[i], values[i + 1]
            if lat == lats[i]:
                return a
            if a is None or b is None:
                return None
            return round(a + (b - a) * (lat - lats[i]) / (lats[i + 1] - lats[i]), 1)
    return None


def solar_declination(date: dt.date):
    """Склонение Солнца в полдень UTC, градусы (алгоритм NOAA / Meeus, точность ~0,01°)."""
    jd = date.toordinal() + 1721424.5 + 0.5
    T = (jd - 2451545) / 36525
    L0 = (280.46646 + T * (36000.76983 + 0.0003032 * T)) % 360
    M = math.radians(357.52911 + T * (35999.05029 - 0.0001537 * T))
    C = ((1.914602 - T * (0.004817 + 0.000014 * T)) * math.sin(M) + (0.019993 - 0.000101 * T) * math.sin(2 * M)
         + 0.000289 * math.sin(3 * M))
    om = math.radians(125.04 - 1934.136 * T)
    lam = math.radians(L0 + C - 0.00569 - 0.00478 * math.sin(om))
    eps = 23 + (26 + (21.448 - T * (46.815 + T * (0.00059 - T * 0.001813))) / 60) / 60 + 0.00256 * math.cos(om)
    return math.degrees(math.asin(math.sin(math.radians(eps)) * math.sin(lam)))


def sun_table(lat, year=2025):
    """Табл. 15.1: склонение, полуденная высота Солнца и долгота дня на 15-е число месяца."""
    out = []
    for m in range(1, 13):
        d = solar_declination(dt.date(year, m, 15))
        h = 90 - lat + d
        h = 180 - h if h > 90 else h
        phi, dr = math.radians(lat), math.radians(d)
        cosw = (math.sin(math.radians(-0.833)) - math.sin(phi) * math.sin(dr)) / (math.cos(phi) * math.cos(dr))
        if cosw <= -1:
            day = 24.0
        elif cosw >= 1:
            day = 0.0
        else:
            day = 2 * math.degrees(math.acos(cosw)) / 15
        out.append((round(d, 2), round(h, 1), day))
    return out


def hhmm(hours):
    m = int(round(hours * 60))
    return f"{m // 60} ч {m % 60:02d} мин"


def sat_pressure_hpa(t):
    """Давление насыщенного пара, гПа (Магнус: надо льдом при t<0, над водой при t≥0)."""
    if t < 0:
        return 6.112 * math.exp(22.46 * t / (272.62 + t))
    return 6.112 * math.exp(17.62 * t / (243.12 + t))


def enthalpy(t, rh, p_hpa):
    """Удельная энтальпия влажного воздуха, кДж/кг сух. возд.; влагосодержание, г/кг."""
    if t is None or rh is None or p_hpa is None:
        return None, None
    e = rh / 100 * sat_pressure_hpa(t)
    d = 622 * e / (p_hpa - e)  # г/кг
    h = 1.006 * t + d / 1000 * (2501 + 1.86 * t)
    return round(h, 1), round(d, 2)


def frost_depth(mt, d0, kh=None):
    if mt is None:
        return None, None
    dfn = d0 * math.sqrt(mt)
    df = dfn * kh if kh is not None else None
    return round(dfn, 2), (round(df, 2) if df is not None else None)


def nearest_osr(osr: pd.DataFrame, lat, lon):
    d = haversine_km(lat, lon, osr["lat"].values, osr["lon"].values)
    i = int(np.argmin(d))
    return osr.iloc[i], float(d[i])


def nearest_stations(st: pd.DataFrame, lat, lon, n=5):
    d = haversine_km(lat, lon, st["lat"].values, st["lon"].values)
    out = st.assign(dist_km=np.round(d, 1)).sort_values("dist_km").head(n)
    return out


# ----------------------------------------------------------------- разделы
@dataclass
class Section:
    title: str
    df: pd.DataFrame
    note: str = ""
    ref: list[str] = field(default_factory=list)  # колонки «справочно» (для серого цвета)


def _param_table(r, prev_rows, spec, with_prev):
    rows = []
    for key, n, name, unit in spec:
        row = {"Графа": n, "Параметр": name, "Ед.": unit, CUR: r.get(key)}
        if with_prev:
            for ed in PREV_EDITIONS:
                pr = prev_rows.get(ed)
                row[ed] = pr.get(key) if pr is not None else None
        rows.append(row)
    df = pd.DataFrame(rows)
    if with_prev:
        df = df[[c for c in df.columns if c in ("Графа", "Параметр", "Ед.", CUR) or df[c].notna().any()]]
    return df


def _monthly_rows(r, prev_rows, pfx, label, with_prev, with_year=True):
    cols = MONTHS + (["Год"] if with_year else [])

    def row_of(src, name):
        d = {"Источник": name}
        for m in MONTHS:
            d[m] = src.get(f"{pfx}_{m}")
        if with_year:
            d["Год"] = src.get(f"{pfx}_year")
        return d

    rows = [row_of(r, CUR)]
    if with_prev:
        for ed in PREV_EDITIONS:
            pr = prev_rows.get(ed)
            if pr is not None and any(pr.get(f"{pfx}_{m}") is not None for m in MONTHS):
                rows.append(row_of(pr, ed))
    return pd.DataFrame(rows, columns=["Источник"] + cols)


def build_sections(r: dict, prev_rows: dict, with_prev=True, soil="суглинки и глины", kh=None,
                   osr_point=None, kh_label="") -> list[Section]:
    S: list[Section] = []
    ref_cols = PREV_EDITIONS if with_prev else []

    # 0. Общие сведения
    info = [
        ("Пункт (метеостанция)", r.get("name")),
        ("Субъект РФ", r.get("region")),
        ("Полное наименование", r.get("address")),
        ("Широта, °", f"{r['lat']:.6f}" if r.get("lat") is not None else None),
        ("Долгота, °", f"{r['lon']:.6f}" if r.get("lon") is not None else None),
        ("№ пункта в таблицах СП 131.13330.2025", r.get("id")),
    ]
    S.append(Section("Общие сведения", pd.DataFrame(info, columns=["Показатель", "Значение"])))

    # 5.1 / 6.1
    S.append(Section("Таблица 5.1 — Климатические параметры холодного периода года",
                     _param_table(r, prev_rows, COLD, with_prev), ref=ref_cols))
    S.append(Section("Таблица 6.1 — Климатические параметры теплого периода года",
                     _param_table(r, prev_rows, WARM, with_prev), ref=ref_cols))

    # 7.1
    S.append(Section("Таблица 7.1 — Средняя месячная и годовая температура воздуха, °С",
                     _monthly_rows(r, prev_rows, "t", "", with_prev)))

    # 8.1
    if r.get("t81_ampl_jul") is not None:
        S.append(Section("Таблица 8.1 — Максимальная суточная амплитуда температуры воздуха в июле, °С",
                         pd.DataFrame([{"Пункт по табл. 8.1": r.get("t81_name"),
                                        "Амплитуда, °С": r.get("t81_ampl_jul")}])))

    # 9.1
    av, mx = {"Показатель": "Средняя"}, {"Показатель": "Максимальная"}
    any91 = False
    for m in MONTHS:
        s = r.get(f"ampl_{m}")
        a = b = None
        if s:
            parts = str(s).split("/")
            try:
                a = float(parts[0].replace(",", "."))
                b = float(parts[1].replace(",", ".")) if len(parts) > 1 else None
                any91 = True
            except ValueError:
                pass
        av[m], mx[m] = a, b
    if any91:
        S.append(Section("Таблица 9.1 — Средняя и максимальная суточная амплитуда температуры наружного воздуха, °С",
                         pd.DataFrame([av, mx], columns=["Показатель"] + MONTHS)))

    # 10.1, 11.1, 12.1
    S.append(Section("Таблица 10.1 — Средняя месячная и годовая сумма осадков, мм",
                     _monthly_rows(r, prev_rows, "pr", "", False)))
    S.append(Section("Таблица 11.1 — Среднее месячное и годовое парциальное давление водяного пара, гПа",
                     _monthly_rows(r, prev_rows, "e", "", with_prev)))
    S.append(Section("Таблица 12.1 — Средняя месячная и годовая скорость ветра, м/с",
                     _monthly_rows(r, prev_rows, "v", "", False)))

    # 13.1, 13.2 — солнечная радиация (нормативные таблицы по широтам, интерполяция на широту станции)
    lat = r.get("lat")
    t1 = REF.get("t13_1")
    t2 = REF.get("t13_2")
    lat_note = (f"Значения интерполированы линейно на широту станции {lat:.2f}° с.ш. по таблице СП "
                f"(широты {t1['lats'][0]:g}–{t1['lats'][-1]:g}° с.ш.). Вне этого диапазона значения не приводятся.")
    if t1:
        row = {"Поверхность": "Горизонтальная"}
        for i, m in enumerate(MONTHS):
            row[m] = _interp_lat(t1["lats"], t1["rows"][i], lat)
        vals = [row[m] for m in MONTHS]
        row["Сумма"] = round(sum(vals), 1) if all(v is not None for v in vals) else None
        S.append(Section("Таблица 13.1 — Суммарная солнечная радиация (прямая и рассеянная) на горизонтальную "
                         "поверхность при безоблачном небе, МДж/м²",
                         pd.DataFrame([row], columns=["Поверхность"] + MONTHS + ["Сумма"]), note=lat_note))
    if t2:
        rows = []
        for k, o in enumerate(t2["orients"]):
            row = {"Ориентация": o}
            for i, m in enumerate(MONTHS):
                row[m] = _interp_lat(t2["lats"], t2["blocks"][i][k], lat)
            rows.append(row)
        S.append(Section("Таблица 13.2 — Суммарная солнечная радиация (прямая и рассеянная) на вертикальную "
                         "поверхность при безоблачном небе, МДж/м²",
                         pd.DataFrame(rows, columns=["Ориентация"] + MONTHS), note=lat_note))

    # 14
    if r.get("t14"):
        rows = []
        for m, chunk in zip(MONTHS, str(r["t14"]).split(";")):
            vals = chunk.split("|")
            rows.append([m] + [v if v else "—" for v in vals])
        S.append(Section("Таблица 14 — Суточный ход рассеянной (числитель) и суммарной (знаменатель) "
                         "освещенности под безоблачным небом, клк",
                         pd.DataFrame(rows, columns=["Месяц"] + [str(i) for i in range(1, 13)]),
                         note="Столбцы 1–12 — графы суточного хода по порядку таблицы 14 СП."))

    # 15.1 — высота Солнца (астрономический расчет на широту станции)
    if lat is not None:
        sun = sun_table(lat)
        S.append(Section("Таблица 15.1 — Значения высоты солнца над горизонтом, градусы",
                         pd.DataFrame([
                             ["Склонение Солнца на 15-е число, °"] + [s[0] for s in sun],
                             ["Высота Солнца в полдень на 15-е число, °"] + [s[1] for s in sun],
                             ["Долгота дня на 15-е число"] + [hhmm(s[2]) for s in sun],
                         ], columns=["Параметр"] + MONTHS),
                         note=f"Астрономический расчет для широты {lat:.4f}° с.ш. (2025 г.): высота = 90° − φ + δ; "
                              "долгота дня — по верхнему краю диска с учетом рефракции (−0,833°)."))

    # 16.2
    if r.get("t162_h") is not None:
        S.append(Section("Таблица 16.2 — Удельная энтальпия и влагосодержание наружного воздуха",
                         pd.DataFrame([
                             {"Параметр": "Удельная энтальпия наружного воздуха обеспеченностью 0,98",
                              "Ед.": "кДж/кг", "Значение": r.get("t162_h")},
                             {"Параметр": "Удельное влагосодержание наружного воздуха",
                              "Ед.": "г/кг", "Значение": r.get("t162_d")},
                         ])))

    # Раздел 16 — параметры для ОВиК
    p = r.get("w_p")
    wmin = r.get("w_wmin")
    wwarm = max(wmin, 1.0) if wmin is not None else None
    wcold = max(r["c_wmax"], 1.0) if r.get("c_wmax") is not None else None
    hA, _ = enthalpy(r.get("c_t094"), r.get("c_rh"), p)
    hB, _ = enthalpy(r.get("c_5d092"), r.get("c_rh"), p)
    hvac = pd.DataFrame([
        {"Период": "Теплый", "Параметры": "А", "t, °С": r.get("w_t095"),
         "Энтальпия, кДж/кг": "по рис. А.5, А.7", "Скорость ветра, м/с": wwarm,
         "Давление, гПа": p, "Ср. сут. амплитуда, °С": r.get("w_ampl")},
        {"Период": "Теплый", "Параметры": "Б", "t, °С": r.get("w_t098"),
         "Энтальпия, кДж/кг": r.get("t162_h") if r.get("t162_h") is not None else "по рис. А.6, А.8",
         "Скорость ветра, м/с": wwarm, "Давление, гПа": p, "Ср. сут. амплитуда, °С": r.get("w_ampl")},
        {"Период": "Холодный", "Параметры": "А", "t, °С": r.get("c_t094"),
         "Энтальпия, кДж/кг": hA, "Скорость ветра, м/с": wcold,
         "Давление, гПа": "не нормируется", "Ср. сут. амплитуда, °С": r.get("c_ampl")},
        {"Период": "Холодный", "Параметры": "Б", "t, °С": r.get("c_5d092"),
         "Энтальпия, кДж/кг": hB, "Скорость ветра, м/с": wcold,
         "Давление, гПа": "не нормируется", "Ср. сут. амплитуда, °С": r.get("c_ampl")},
    ])
    S.append(Section("Таблица 16.1 — Климатические параметры для проектирования отопления, вентиляции, "
                     "кондиционирования и холодоснабжения", hvac,
                     note="Теплый период: А — t обесп. 0,95, Б — t обесп. 0,98 (табл. 6.1); скорость ветра — "
                          "минимальная из средних по румбам за июль, не менее 1 м/с. Холодный период: А — t обесп. 0,94, "
                          "Б — t наиболее холодной пятидневки обесп. 0,92 (табл. 5.1); скорость ветра — максимальная из "
                          "средних по румбам за январь, не менее 1 м/с. Энтальпия холодного периода рассчитана по t и средней месячной "
                          "относительной влажности наиболее холодного месяца (табл. 5.1, графа 15) при барометрическом "
                          "давлении по табл. 6.1."))

    # Приложение А, табл. А.1 — климатическое районирование
    if REF.get("tA1"):
        subs = set(climate_subregions(r))
        vw_ = [r.get("v_XII"), r.get("v_I"), r.get("v_II")]
        w3 = round(sum(vw_) / 3, 1) if all(v is not None for v in vw_) else None
        a1 = [[reg_, sub, tj, w, tjul, rh, "да" if sub in subs else ""]
              for reg_, sub, tj, w, tjul, rh in REF["tA1"]["rows"]]
        a1.append(["Станция", "", r.get("t_I"), w3, r.get("t_VII"),
                   f"{fmt(r.get('w_rh'))} (в 15 ч — {fmt(r.get('w_rh15'))})", ""])
        S.append(Section("Таблица А.1 — Климатическое районирование для строительства",
                         pd.DataFrame(a1, columns=["Район", "Подрайон", "t января, °С",
                                                   "Ветер за 3 зимних мес., м/с", "t июля, °С",
                                                   "Отн. влажность в июле, %", "Условия выполнены"]),
                         note=(REF["tA1"].get("note") or "") + ". Последняя строка — данные станции "
                              "(t января и июля — табл. 7.1; ветер — среднее за XII–II по табл. 12.1; влажность — "
                              "наиболее теплого месяца по табл. 6.1, гр. 8–9; продолжительность периода ≤ 0 °С — "
                              f"{fmt(r.get('c_n0'))} сут). Отметка «да» — проверка приложения, окончательно — по "
                              "карте рис. А.1."))

    # Приложение Б, табл. Б.1 — подзоны суровых условий
    if REF.get("tB1"):
        vals, zones = severity_check(r)
        cols = ["Подзона", "Строка", "Абс. мин., °С", "Сутки 0,98, °С", "Сутки 0,92, °С",
                "Пятидневка 0,98, °С", "Пятидневка 0,92, °С", "Сумма t при t < 8 °С, °С·сут"]
        b1 = []
        for i, row_ in enumerate(REF["tB1"]["rows"]):
            b1.append([row_[0], "макс." if i % 2 == 0 else "мин."] + row_[1:])
        b1.append(["Станция", ""] + vals)
        for name, ok in zones.items():
            b1.append([f"Попадание: {name.lower()}", ""] + ["да" if x else "" for x in ok])
        S.append(Section("Таблица Б.1 — Подзоны суровых условий", pd.DataFrame(b1, columns=cols),
                         note="Первая строка подзоны — максимальные, вторая — минимальные значения. "
                              + (REF["tB1"].get("note") or "") + " (для станции — произведение граф 11 и 12 "
                              "табл. 5.1). Строки «Попадание» — в диапазон какой подзоны попадает каждый показатель "
                              "станции; подзона устанавливается для северной строительно-климатической зоны (рис. А.2)."))

    # Производные: климат. подрайон, промерзание
    mt = r.get("mt")
    d0 = SOILS[soil]
    dfn, df = frost_depth(mt, d0, kh)
    vw = [r.get("v_XII"), r.get("v_I"), r.get("v_II")]
    wind_w = round(sum(vw) / 3, 1) if all(v is not None for v in vw) else None
    derived = [
        ("Климатический подрайон — условия табл. А.1 выполняются для", r.get("climate_sub") or "не определяется",
         "условно; окончательно — по карте рис. А.1"),
        ("Средняя месячная температура января", r.get("t_I"), "°С"),
        ("Средняя месячная температура июля", r.get("t_VII"), "°С"),
        ("Средняя скорость ветра за декабрь–февраль (табл. 12.1)", wind_w, "м/с"),
        ("Мt — сумма абсолютных значений среднемесячных отрицательных температур", mt, "°С"),
        ("Средняя из отрицательных среднемесячных температур (для п. 6.8.3 СП 22.13330)", winter_mean_negative(r), "°С"),
        (f"d0 для грунта «{soil}»", d0, "м"),
        ("Нормативная глубина сезонного промерзания dfn = d0·√Mt (п. 5.5.3 СП 22.13330)", dfn, "м"),
    ]
    if kh is not None:
        derived.append((f"Расчетная глубина промерзания df = kh·dfn при kh = {fmt(kh)} (п. 5.5.4)", df,
                        "м" + (f"; kh по табл. 5.2: {kh_label}" if kh_label else "")))
    note = ""
    if mt == 0:
        note = "Отрицательных среднемесячных температур нет (Mt = 0): сезонное промерзание по формуле п. 5.5.3 не определяется."
    if dfn is not None and dfn > 2.5:
        note = "dfn > 2,5 м: по п. 5.5.3 СП 22.13330 глубину промерзания следует определять теплотехническим расчетом."
    S.append(Section("Производные параметры (расчет по данным СП 131.13330.2025)",
                     pd.DataFrame(derived, columns=["Параметр", "Значение", "Ед. / примечание"]), note=note))

    # Справочно: снег (СП 20) и сейсмичность (ОСР-2016)
    ref = []
    if r.get("sp20_name"):
        ref.append(("Вес снегового покрова Sg, СП 20.13330.2016, Прил. К",
                    r.get("sg_kpa") if r.get("sg_kpa") is not None else "в Прил. К не указан — по карте 1 СП 20",
                    "кПа" if r.get("sg_kpa") is not None else "",
                    f"{r.get('sp20_name')} ({fmt(r.get('sp20_dist_km'))} км от станции)"))
    else:
        ref.append(("Вес снегового покрова Sg, СП 20.13330.2016, Прил. К", "пункт в Прил. К не найден",
                    "", "определять по карте 1 СП 20"))
    if r.get("osr_name"):
        ref.append(("Сейсмичность по ОСР-2016 A / B / C (у станции)",
                    f"{r.get('osr_A')} / {r.get('osr_B')} / {r.get('osr_C')}", "балл MSK-64",
                    f"{r.get('osr_name')} ({fmt(r.get('osr_dist_km'))} км от станции)"))
        if any(r.get(k) for k in ("osr_A_2015", "osr_B_2015", "osr_C_2015")):
            ref.append(("То же по ОСР-2015 (в источнике — в скобках)",
                        f"{r.get('osr_A_2015') or '—'} / {r.get('osr_B_2015') or '—'} / {r.get('osr_C_2015') or '—'}",
                        "балл MSK-64", ""))
    if osr_point is not None:
        o, dist, (plat, plon) = osr_point
        ref.append((f"Сейсмичность по ОСР-2016 A / B / C (для точки {plat:.5f}, {plon:.5f})",
                    f"{o['A']} / {o['B']} / {o['C']}", "балл MSK-64",
                    f"{o['full_name'] or o['name']} ({dist:.1f} км от точки)"))
    S.append(Section("Справочно: снеговая нагрузка и сейсмичность",
                     pd.DataFrame(ref, columns=["Параметр", "Значение", "Ед.", "Пункт / примечание"]),
                     note="Сейсмичность — по ближайшему населенному пункту перечня Прил. А СП 14.13330.2018 (изм. 1) "
                          "с баллами по картам ОСР-2016. Для площадки принимать по перечню СП 14 и заданию."))
    return S


# ----------------------------------------------------------------- экспорт
def _cell(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (np.floating,)):
        return float(v)
    if isinstance(v, (np.integer,)):
        return int(v)
    return v


def to_excel(sections: list[Section], station_title: str) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Климат"
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="E7E9EC")
    row = 1
    ws.cell(row, 1, f"Климатические параметры: {station_title}").font = Font(bold=True, size=13)
    row += 1
    ws.cell(row, 1, SOURCES).alignment = Alignment(wrap_text=False)
    row += 2
    maxw = {}
    for s in sections:
        ws.cell(row, 1, s.title).font = Font(bold=True, size=11)
        row += 1
        for j, c in enumerate(s.df.columns, 1):
            cell = ws.cell(row, j, str(c))
            cell.font = Font(bold=True)
            cell.fill = hdr_fill
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            maxw[j] = max(maxw.get(j, 0), min(len(str(c)), 18))
        row += 1
        for rec in s.df.itertuples(index=False):
            for j, v in enumerate(rec, 1):
                v = _cell(v)
                cell = ws.cell(row, j, v if v is not None else "—")
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if s.df.columns[j - 1] in s.ref:
                    cell.font = Font(color="666666")
                maxw[j] = max(maxw.get(j, 0), min(len(str(v)) if v is not None else 1, 70))
            row += 1
        if s.note:
            ws.cell(row, 1, "Примечание: " + s.note).font = Font(italic=True, size=9)
            row += 1
        row += 1
    for j, w in maxw.items():
        ws.column_dimensions[get_column_letter(j)].width = max(8, w + 2)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def to_word(sections: list[Section], station_title: str) -> bytes:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(1.5))
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(9)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    h = doc.add_paragraph()
    run = h.add_run(f"Климатические параметры: {station_title}")
    run.bold = True
    run.font.size = Pt(13)
    p = doc.add_paragraph(SOURCES)
    p.runs[0].font.size = Pt(8)

    for s in sections:
        t = doc.add_paragraph()
        r = t.add_run(s.title)
        r.bold = True
        r.font.size = Pt(10)
        cols = list(s.df.columns)
        table = doc.add_table(rows=1, cols=len(cols))
        table.style = "Table Grid"
        for j, c in enumerate(cols):
            cell = table.rows[0].cells[j]
            cell.text = str(c)
            cell.paragraphs[0].runs[0].bold = True
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:fill"), "E7E9EC")
            cell._tc.get_or_add_tcPr().append(shd)
        for rec in s.df.itertuples(index=False):
            cells = table.add_row().cells
            for j, v in enumerate(rec):
                cells[j].text = fmt(_cell(v))
                if cols[j] in s.ref:
                    for run_ in cells[j].paragraphs[0].runs:
                        run_.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        for row_ in table.rows:
            for c in row_.cells:
                for par in c.paragraphs:
                    for run_ in par.runs:
                        run_.font.size = Pt(8)
        if "Параметр" in cols:
            j = cols.index("Параметр")
            for row_ in table.rows:
                row_.cells[j].width = Cm(9)
        if s.note:
            n = doc.add_paragraph("Примечание: " + s.note)
            n.runs[0].italic = True
            n.runs[0].font.size = Pt(8)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def list_to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Станции")
        ws = xw.sheets["Станции"]
        from openpyxl.utils import get_column_letter
        for j, c in enumerate(df.columns, 1):
            w = max([len(str(c))] + [len(fmt(v)) for v in df[c].head(300)])
            ws.column_dimensions[get_column_letter(j)].width = min(max(8, w + 2), 60)
    return buf.getvalue()
