"""
Сборка данных приложения из исходных xlsb.

Запуск (из папки проекта):
    python tools/build_data.py "<СП131...2025.xlsb>" "<СП20.xlsb>" "<ОСР-2016.xlsb>"

Нужны пакеты: pyxlsb, pandas, msoffcrypto-tool (файл ОСР закрыт стандартным
ключом Excel без пароля — расшифровывается автоматически).

Результат — CSV в папке data/ (их и читает app.py).
"""
import io
import math
import re
import sys
from pathlib import Path

import pandas as pd
from pyxlsb import open_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data"

MONTHS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]


# ----------------------------------------------------------------- helpers
def read_sheet(src, sheet):
    """Весь лист -> {row: {col: value}} (пустые ячейки пропускаются)."""
    data = {}
    with open_workbook(src) as wb:
        with wb.get_sheet(sheet) as sh:
            for i, row in enumerate(sh.rows()):
                d = {c.c: c.v for c in row if c.v not in (None, "")}
                if d:
                    data[i] = d
    return data


def clean(v):
    """Ошибки Excel ('0x..'), прочерки и 'н/д' -> None; строки -> strip."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "-", "–", "—", "н/д") or re.fullmatch(r"0x[0-9a-fA-F]+", s):
            return None
        return s
    if isinstance(v, float) and math.isnan(v):
        return None
    return v


def num(v):
    v = clean(v)
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ".").replace("−", "-"))
    except ValueError:
        return None


def parse_coords(s):
    s = clean(s)
    if not isinstance(s, str):
        return None, None
    m = re.match(r"\s*(-?\d+(?:[.,]\d+)?)\s*,\s*(-?\d+(?:[.,]\d+)?)", s)
    if not m:
        return None, None
    return float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", "."))


def haversine(lat1, lon1, lat2, lon2):
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def norm_name(s):
    s = (s or "").lower().replace("ё", "е")
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^a-zа-я0-9]", "", s)
    return s


# --------------------------------------------- описание параметров (СП131-2025)
# key, колонка листа «Сводная», раздел, наименование, ед.
COLD = [
    ("c_day098", 8, "Температура воздуха наиболее холодных суток, обеспеченностью 0,98", "°С"),
    ("c_day092", 9, "Температура воздуха наиболее холодных суток, обеспеченностью 0,92", "°С"),
    ("c_5d098", 10, "Температура воздуха наиболее холодной пятидневки, обеспеченностью 0,98", "°С"),
    ("c_5d092", 11, "Температура воздуха наиболее холодной пятидневки, обеспеченностью 0,92", "°С"),
    ("c_t094", 12, "Температура воздуха обеспеченностью 0,94", "°С"),
    ("c_absmin", 13, "Абсолютная минимальная температура воздуха", "°С"),
    ("c_ampl", 14, "Средняя суточная амплитуда температуры воздуха наиболее холодного месяца", "°С"),
    ("c_n0", 15, "Продолжительность периода со средней суточной температурой ≤ 0 °С", "сут"),
    ("c_t0", 16, "Средняя температура периода со средней суточной температурой ≤ 0 °С", "°С"),
    ("c_n8", 17, "Продолжительность периода со средней суточной температурой ≤ 8 °С", "сут"),
    ("c_t8", 18, "Средняя температура периода со средней суточной температурой ≤ 8 °С", "°С"),
    ("c_n10", 19, "Продолжительность периода со средней суточной температурой ≤ 10 °С", "сут"),
    ("c_t10", 20, "Средняя температура периода со средней суточной температурой ≤ 10 °С", "°С"),
    ("c_rh", 21, "Средняя месячная относительная влажность воздуха наиболее холодного месяца", "%"),
    ("c_rh15", 22, "Средняя месячная относительная влажность воздуха в 15 ч наиболее холодного месяца", "%"),
    ("c_prec", 23, "Количество осадков за ноябрь–март", "мм"),
    ("c_wdir", 24, "Преобладающее направление ветра за декабрь–февраль", "—"),
    ("c_wmax", 25, "Максимальная из средних скоростей ветра по румбам за январь", "м/с"),
    ("c_wavg", 26, "Средняя скорость ветра за период со средней суточной температурой ≤ 8 °С", "м/с"),
]
WARM = [
    ("w_p", 29, "Барометрическое давление", "гПа"),
    ("w_t095", 30, "Температура воздуха обеспеченностью 0,95", "°С"),
    ("w_t098", 31, "Температура воздуха обеспеченностью 0,98", "°С"),
    ("w_tmax_avg", 32, "Средняя максимальная температура воздуха наиболее теплого месяца", "°С"),
    ("w_absmax", 33, "Абсолютная максимальная температура воздуха", "°С"),
    ("w_ampl", 34, "Средняя суточная амплитуда температуры воздуха наиболее теплого месяца", "°С"),
    ("w_rh", 35, "Средняя месячная относительная влажность воздуха наиболее теплого месяца", "%"),
    ("w_rh15", 36, "Средняя месячная относительная влажность воздуха в 15 ч наиболее теплого месяца", "%"),
    ("w_prec", 37, "Количество осадков за апрель–октябрь", "мм"),
    ("w_precmax", 38, "Суточный максимум осадков", "мм"),
    ("w_wdir", 39, "Преобладающее направление ветра за июнь–август", "—"),
    ("w_wmin", 40, "Минимальная из средних скоростей ветра по румбам за июль", "м/с"),
]
MONTHLY = {  # префикс: (первая колонка I..XII, колонка «год» или None)
    "t": (43, 55),      # табл. 7.1 температура
    "ampl": (61, None),  # табл. 9.1 «средняя/максимальная» (строка)
    "pr": (75, 87),     # табл. 10.1 осадки
    "e": (90, 102),     # табл. 11.1 парциальное давление
    "v": (105, 117),    # табл. 12.1 скорость ветра
}

# предыдущие редакции в листе «Sum»: те же параметры, свои колонки
PREV = {
    "СНиП 23-01-99*": {"cold": 110, "warm": 131, "t": 95, "e": 145, "name": 109},
    "СП 131.13330.2018": {"cold": 239, "warm": 263, "t": 280, "e": 303, "name": 238},
    "СП 131.13330.2020": {"cold": 447, "warm": 467, "t": 480, "e": 496, "name": 445},
}


def build_sp131(src):
    sv = read_sheet(src, "Сводная")
    rows = sorted(r for r in sv if r >= 7 and isinstance(sv[r].get(2), (int, float)))
    recs = []
    for r in rows:
        d = sv[r]
        lat, lon = parse_coords(d.get(1))
        rec = {
            "id": int(d[2]),
            "name": clean(d.get(3)),
            "address": clean(d.get(5)),
            "region": clean(d.get(6)),
            "lat": lat,
            "lon": lon,
            "src_row": r + 1,  # номер строки в Excel (1-based)
        }
        for key, col, *_ in COLD + WARM:
            v = clean(d.get(col))
            rec[key] = v if key.endswith("wdir") else num(v)
        for pfx, (c0, cy) in MONTHLY.items():
            for i, m in enumerate(MONTHS):
                v = clean(d.get(c0 + i))
                rec[f"{pfx}_{m}"] = (str(v).replace(" ", "") if v is not None else None) if pfx == "ampl" else num(v)
            if cy is not None:
                rec[f"{pfx}_year"] = num(d.get(cy))
        # табл. 8.1 — максимальная суточная амплитуда в июле
        rec["t81_name"] = clean(d.get(57))
        rec["t81_ampl_jul"] = num(d.get(58))
        # табл. 16.2 — энтальпия / влагосодержание
        rec["t162_name"] = clean(d.get(277))
        rec["t162_h"] = num(d.get(278))
        rec["t162_d"] = num(d.get(279))
        # табл. 14 — освещенность: 12 месяцев x (метка + 12 значений)
        t14 = []
        for mi in range(12):
            base = 119 + mi * 13
            vals = [clean(d.get(base + 1 + k)) for k in range(12)]
            vals = [str(v).replace(" ", "") if v is not None else "" for v in vals]
            t14.append("|".join(vals))
        rec["t14"] = ";".join(t14) if any(x.replace("|", "") for x in t14) else None
        recs.append(rec)
    df = pd.DataFrame(recs)
    return df, sv


def build_prev(src, st):
    """Предыдущие редакции: сопоставление со станциями 2025 по координатам."""
    sm = read_sheet(src, "Sum")
    rows = [r for r in sorted(sm) if r >= 5]
    sum_pts = []
    for r in rows:
        d = sm[r]
        lat, lon = num(d.get(433)), num(d.get(434))
        if lat is None:
            lat, lon = parse_coords(d.get(435))
        sum_pts.append((r, lat, lon, clean(d.get(428))))
    out = []
    unmatched = []
    for _, s in st.iterrows():
        best, bd = None, 1e9
        for r, lat, lon, nm in sum_pts:
            if lat is None:
                continue
            dist = haversine(s.lat, s.lon, lat, lon)
            if dist < bd:
                best, bd = (r, nm), dist
        if best is None or bd > 3.0:
            unmatched.append(s["name"])
            continue
        d = sm[best[0]]
        for ed, cols in PREV.items():
            rec = {"id": s["id"], "edition": ed, "prev_name": clean(d.get(cols["name"]))}
            has = False
            for i, (key, *_r) in enumerate(COLD):
                v = clean(d.get(cols["cold"] + i))
                rec[key] = v if key.endswith("wdir") else num(v)
                has |= rec[key] is not None
            for i, (key, *_r) in enumerate(WARM):
                v = clean(d.get(cols["warm"] + i))
                rec[key] = v if key.endswith("wdir") else num(v)
                has |= rec[key] is not None
            for i, m in enumerate(MONTHS):
                rec[f"t_{m}"] = num(d.get(cols["t"] + i))
                rec[f"e_{m}"] = num(d.get(cols["e"] + i))
            rec["t_year"] = num(d.get(cols["t"] + 12))
            rec["e_year"] = num(d.get(cols["e"] + 12))
            if has:
                out.append(rec)
    return pd.DataFrame(out), unmatched


def build_sp20(src, st):
    """СП 20.13330.2016, Прил. К — вес снегового покрова Sg по населённым пунктам."""
    l1 = read_sheet(src, "Лист1")
    pts = []
    for r in sorted(l1):
        if r < 4:
            continue
        d = l1[r]
        name = clean(d.get(2))
        lat, lon = parse_coords(d.get(4))
        if not name or lat is None:
            continue
        sg = num(d.get(3))
        pts.append({"sp20_name": name, "sp20_lat": lat, "sp20_lon": lon, "sg_kpa": sg,
                    "sg_note": None if sg is not None else clean(d.get(3))})
    pts_df = pd.DataFrame(pts)
    res = []
    for _, s in st.iterrows():
        dd = pts_df.apply(lambda p: haversine(s.lat, s.lon, p.sp20_lat, p.sp20_lon), axis=1)
        i = dd.idxmin()
        # дубли пункта: при равном удалении предпочесть запись со значением Sg
        near = dd[(dd <= dd[i] + 0.5)]
        with_val = [k for k in near.index if pd.notna(pts_df.loc[k, "sg_kpa"])]
        if with_val:
            i = min(with_val, key=lambda k: dd[k])
        if dd[i] <= 5.0:
            p = pts_df.loc[i]
            res.append({"id": s["id"], "sp20_name": p.sp20_name, "sg_kpa": p.sg_kpa,
                        "sg_note": p.sg_note, "sp20_dist_km": round(dd[i], 2)})
    return pts_df, pd.DataFrame(res)


def open_maybe_encrypted(path):
    with open(path, "rb") as f:
        head = f.read(8)
    if head[:4] == b"PK\x03\x04":
        return path
    import msoffcrypto  # файл закрыт стандартным ключом Excel (без пароля)
    of = msoffcrypto.OfficeFile(open(path, "rb"))
    of.load_key(password="VelvetSweatshop")
    buf = io.BytesIO()
    of.decrypt(buf)
    tmp = OUT / "_osr_tmp.xlsb"
    tmp.write_bytes(buf.getvalue())
    return tmp


def parse_ball(s):
    """'7 (8)' -> (7, 8); '8 (-)' -> (8, None); '<6' сохраняется как текст."""
    s = clean(s)
    if s is None:
        return None, None
    s = str(s)
    m = re.match(r"\s*([^()]+?)\s*(?:\((.*?)\))?\s*$", s)
    a = m.group(1).strip() if m else s
    b = (m.group(2) or "").strip() if m else ""
    return a, (b if b not in ("", "-") else None)


def build_osr(src):
    path = open_maybe_encrypted(src)
    try:
        sh = read_sheet(path, "Прил.А СП14(2018) изм.1+")
    finally:
        if Path(path).name == "_osr_tmp.xlsb":
            Path(path).unlink(missing_ok=True)
    recs = []
    for r in sorted(sh):
        if r < 2:
            continue
        d = sh[r]
        lat, lon = parse_coords(d.get(4))
        if lat is None:
            continue
        a, a15 = parse_ball(d.get(5))
        b, b15 = parse_ball(d.get(6))
        c, c15 = parse_ball(d.get(7))
        recs.append({
            "region": (clean(d.get(1)) or "").replace("\n", " "),
            "name": clean(d.get(2)),
            "full_name": clean(d.get(3)),
            "lat": lat, "lon": lon,
            "A": a, "B": b, "C": c,
            "A_2015": a15, "B_2015": b15, "C_2015": c15,
        })
    return pd.DataFrame(recs)


def nearest_osr(st, osr):
    import numpy as np
    la = np.radians(osr.lat.values)
    lo = np.radians(osr.lon.values)
    res = []
    for _, s in st.iterrows():
        p1, l1 = math.radians(s.lat), math.radians(s.lon)
        a = np.sin((la - p1) / 2) ** 2 + np.cos(p1) * np.cos(la) * np.sin((lo - l1) / 2) ** 2
        dist = 2 * 6371 * np.arcsin(np.sqrt(a))
        i = int(dist.argmin())
        o = osr.iloc[i]
        res.append({"id": s["id"], "osr_name": o.full_name or o["name"], "osr_dist_km": round(float(dist[i]), 2),
                    "osr_A": o.A, "osr_B": o.B, "osr_C": o.C,
                    "osr_A_2015": o.A_2015, "osr_B_2015": o.B_2015, "osr_C_2015": o.C_2015})
    return pd.DataFrame(res)


def build_ref_tables(src):
    """Нормативные справочные таблицы, не привязанные к станции:
    13.1 и 13.2 (по широтам), А.1, Б.1 (СП 131), табл. 5.2 СП 22 (kh)."""
    calc = read_sheet(src, "Расчетка")
    main_ = read_sheet(src, "СП131.13330.2025")
    ref = {}

    # 13.1 — горизонтальная поверхность: строки 78..89 (I..XII), широты в строке 77, кол. 10..17
    lats = [num(calc[77][c]) for c in range(10, 18)]
    ref["t13_1"] = {"lats": lats,
                    "rows": [[num(calc.get(78 + i, {}).get(c)) for c in range(10, 18)] for i in range(12)]}
    assert [str(calc[78 + i][9]) for i in (0, 11)] == ["Январь", "Декабрь"]

    # 13.2 — вертикальные поверхности: блоки по 6 строк с 64 (месяц + 5 ориентаций), кол. 25..32
    lats2 = [num(calc[63][c]) for c in range(25, 33)]
    orients = ["С", "СВ/СЗ", "В/З", "ЮВ/ЮЗ", "Ю"]
    blocks = []
    for i in range(12):
        r0 = 64 + i * 6
        assert str(calc[r0 + 1][24]).strip() == "С", r0
        blocks.append([[num(calc.get(r0 + 1 + k, {}).get(c)) for c in range(25, 33)] for k in range(5)])
        for k, o in enumerate(orients):
            assert str(calc[r0 + 1 + k][24]).strip() == o
    ref["t13_2"] = {"lats": lats2, "orients": orients, "blocks": blocks}

    # А.1 — строки 42..57; примечание — строка 58
    a1 = []
    region = None
    for r in range(42, 58):
        d = calc[r]
        region = clean(d.get(1)) or region
        a1.append([region, clean(d.get(2)), clean(d.get(3)), clean(d.get(5)), clean(d.get(7)), clean(d.get(9))])
    ref["tA1"] = {"rows": a1, "note": clean(calc[58].get(1))}

    # Б.1 — строки 67..72 однолистовой формы: (подзона; макс/мин) x 6 показателей
    b1 = []
    for r in range(67, 73):
        d = main_[r]
        name = clean(d.get(1))
        b1.append([name] + [num(d.get(c)) for c in range(2, 8)])
    for i in range(1, 6, 2):  # вторая строка подзоны без названия
        b1[i][0] = b1[i - 1][0]
    ref["tB1"] = {"rows": b1, "note": clean(main_[74].get(1))}

    # СП 22.13330 табл. 5.2 — коэффициент kh
    ref["kh"] = {"temps": [clean(calc[19].get(c)) for c in range(2, 8)],
                 "rows": [[clean(calc[r].get(1))] + [num(calc[r].get(c)) for c in range(2, 8)] for r in range(22, 26)]}
    return ref


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(1)
    sp131, sp20, osr_src = sys.argv[1:4]
    OUT.mkdir(exist_ok=True)

    st, _ = build_sp131(sp131)
    print(f"СП131-2025: станций {len(st)}")

    prev, unmatched = build_prev(sp131, st)
    print(f"Предыдущие редакции: записей {len(prev)}; не сопоставлено станций: {len(unmatched)}")

    sp20_pts, sp20_match = build_sp20(sp20, st)
    print(f"СП20 Прил.К: пунктов {len(sp20_pts)}, сопоставлено со станциями {len(sp20_match)}")

    osr = build_osr(osr_src)
    print(f"ОСР-2016 (перечень СП14): пунктов {len(osr)}")
    osr_near = nearest_osr(st, osr)

    st = st.merge(sp20_match, on="id", how="left").merge(osr_near, on="id", how="left")
    st.to_csv(OUT / "stations.csv", index=False, encoding="utf-8")
    prev.to_csv(OUT / "prev_editions.csv", index=False, encoding="utf-8")
    sp20_pts.to_csv(OUT / "sp20_snow.csv", index=False, encoding="utf-8")
    osr.to_csv(OUT / "osr2016.csv", index=False, encoding="utf-8")

    import json
    ref = build_ref_tables(sp131)
    (OUT / "ref_tables.json").write_text(json.dumps(ref, ensure_ascii=False, indent=1), encoding="utf-8")
    print("Справочные таблицы 13.1, 13.2, А.1, Б.1, kh: записаны")
    print("Готово:", OUT)


if __name__ == "__main__":
    main()

