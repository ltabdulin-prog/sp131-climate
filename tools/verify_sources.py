"""
Перепроверка данных приложения по исходным xlsb.

    python tools/verify_sources.py "<СП131...2025.xlsb>" "<СП20.xlsb>"

1) Каждое значение stations.csv сравнивается с ячейкой листа «Сводная».
2) Значения однолистовой формы «СП131.13330.2025» (выбранный в ней пункт)
   сравниваются с данными приложения — независимая выборка через формулы Excel.
3) Согласованность: годовая сумма осадков = сумма месяцев; год. температура ≈ среднее;
   0,98 ≤ 0,92 для холодного периода; 0,98 ≥ 0,95 для теплого.
4) Сопоставление прежних редакций: совпадение названий.
5) СП20 Прил. К: значение Sg совпадает с листом «Лист1».
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import build_data as b  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
st = pd.read_csv(ROOT / "data" / "stations.csv", encoding="utf-8")
prev = pd.read_csv(ROOT / "data" / "prev_editions.csv", encoding="utf-8")
sp131, sp20 = sys.argv[1], sys.argv[2]
problems = []

# 1 -----------------------------------------------------------------------
sv = b.read_sheet(sp131, "Сводная")
checked = 0
colmap = {k: c for k, c, *_ in b.COLD + b.WARM}
for pfx, (c0, cy) in b.MONTHLY.items():
    for i, m in enumerate(b.MONTHS):
        colmap[f"{pfx}_{m}"] = c0 + i
    if cy:
        colmap[f"{pfx}_year"] = cy
for _, s in st.iterrows():
    row = sv[int(s.src_row) - 1]
    assert int(row[2]) == int(s.id), s.id
    for key, col in colmap.items():
        src = b.clean(row.get(col))
        app = s[key]
        if key.startswith("ampl_") or key.endswith("wdir"):
            a = None if pd.isna(app) else str(app)
            sv_ = None if src is None else (str(src).replace(" ", "") if key.startswith("ampl_") else str(src))
            ok = a == sv_
        else:
            sv_ = b.num(src)
            ok = (pd.isna(app) and sv_ is None) or (sv_ is not None and not pd.isna(app) and abs(sv_ - app) < 1e-9)
        checked += 1
        if not ok:
            problems.append(f"[1] {s['name']}: {key} приложение={app} источник={src}")
print(f"1) Сверено значений со «Сводной»: {checked}, расхождений: {sum(p.startswith('[1]') for p in problems)}")

# 2 -----------------------------------------------------------------------
mn = b.read_sheet(sp131, "СП131.13330.2025")
label = str(mn[2].get(13, ""))
print("2) Однолистовая форма:", label)
name = label.split("станция")[1].split("(")[0].strip()
cand = st[st["name"] == name]
s = cand.iloc[0]
pairs = []
cold_rows = [(4, "c_day098"), (5, "c_day092"), (6, "c_5d098"), (7, "c_5d092"), (8, "c_t094"), (9, "c_absmin"),
             (10, "c_ampl"), (11, "c_n0"), (12, "c_t0"), (13, "c_n8"), (14, "c_t8"), (15, "c_n10"), (16, "c_t10"),
             (17, "c_rh"), (18, "c_rh15"), (19, "c_prec"), (20, "c_wdir"), (21, "c_wmax"), (22, "c_wavg")]
for rr, key in cold_rows:
    pairs.append((key, mn[rr].get(10)))
warm_rows = [(4, "w_p"), (5, "w_t095"), (6, "w_t098"), (7, "w_tmax_avg"), (8, "w_absmax"), (9, "w_ampl"),
             (10, "w_rh"), (11, "w_rh15"), (12, "w_prec"), (13, "w_precmax"), (14, "w_wdir"), (15, "w_wmin")]
for rr, key in warm_rows:
    pairs.append((key, mn[rr].get(17)))
for i, m in enumerate(b.MONTHS):
    pairs.append((f"pr_{m}", mn[26].get(5 + i)))
    pairs.append((f"e_{m}", mn[28].get(5 + i)))
    pairs.append((f"v_{m}", mn[30].get(5 + i)))
pairs += [("pr_year", mn[26].get(17)), ("e_year", mn[28].get(17)), ("v_year", mn[30].get(17))]
bad2 = 0
for key, v in pairs:
    a = s[key]
    if key.endswith("wdir"):
        ok = str(a) == str(v).strip()
    else:
        ok = b.num(v) is not None and abs(b.num(v) - float(a)) < 1e-9
    if not ok:
        bad2 += 1
        problems.append(f"[2] {name}: {key} приложение={a} одностраничная форма={v}")
print(f"   сверено {len(pairs)} значений, расхождений: {bad2}")

# 3 -----------------------------------------------------------------------
bad3 = 0
for _, s in st.iterrows():
    pr = [s[f"pr_{m}"] for m in b.MONTHS]
    if all(pd.notna(pr)) and pd.notna(s.pr_year) and abs(sum(pr) - s.pr_year) > 2:
        bad3 += 1
        problems.append(f"[3] {s['name']}: сумма осадков по месяцам {sum(pr):g} ≠ год {s.pr_year:g}")
    t = [s[f"t_{m}"] for m in b.MONTHS]
    if all(pd.notna(t)) and pd.notna(s.t_year) and abs(sum(t) / 12 - s.t_year) > 0.3:
        bad3 += 1
        problems.append(f"[3] {s['name']}: среднее месячных t {sum(t)/12:.2f} ≠ год {s.t_year:g}")
    if pd.notna(s.c_day098) and pd.notna(s.c_day092) and s.c_day098 > s.c_day092:
        bad3 += 1
        problems.append(f"[3] {s['name']}: t сут 0,98 ({s.c_day098}) > 0,92 ({s.c_day092})")
    if pd.notna(s.c_5d098) and pd.notna(s.c_5d092) and s.c_5d098 > s.c_5d092:
        bad3 += 1
        problems.append(f"[3] {s['name']}: t 5-дн 0,98 ({s.c_5d098}) > 0,92 ({s.c_5d092})")
    if pd.notna(s.w_t098) and pd.notna(s.w_t095) and s.w_t098 < s.w_t095:
        bad3 += 1
        problems.append(f"[3] {s['name']}: t тепл 0,98 ({s.w_t098}) < 0,95 ({s.w_t095})")
    if pd.notna(s.c_absmin) and pd.notna(s.c_day098) and s.c_absmin > s.c_day098:
        bad3 += 1
        problems.append(f"[3] {s['name']}: абс. мин ({s.c_absmin}) выше t сут 0,98 ({s.c_day098})")
print(f"3) Проверки согласованности: замечаний {bad3}")

# 4 -----------------------------------------------------------------------
bad4 = 0
names = st.set_index("id")["name"].to_dict()
for _, p in prev.iterrows():
    pn = b.norm_name(str(p.prev_name).replace("*", "")) if pd.notna(p.prev_name) else ""
    sn = b.norm_name(names[p.id])
    if pn and not (pn.startswith(sn[:5]) or sn.startswith(pn[:5])):
        bad4 += 1
        problems.append(f"[4] {names[p.id]} ↔ {p.edition}: «{p.prev_name}»")
print(f"4) Прежние редакции: записей {len(prev)}, станций {prev.id.nunique()}, "
      f"несовпадений названий (на проверку): {bad4}")

# 5 -----------------------------------------------------------------------
l1 = b.read_sheet(sp20, "Лист1")
byname = {}
for r, d in l1.items():  # при дублях названия берём запись со значением Sg
    if r >= 4 and (b.clean(d.get(2)) not in byname or b.num(d.get(3)) is not None):
        byname[b.clean(d.get(2))] = d
bad5 = n5 = 0
for _, s in st[st.sp20_name.notna()].iterrows():
    d = byname.get(s.sp20_name)
    n5 += 1
    v = b.num(d.get(3)) if d else None
    if not ((pd.isna(s.sg_kpa) and v is None) or (v is not None and abs(v - s.sg_kpa) < 1e-9)):
        bad5 += 1
        problems.append(f"[5] {s['name']}: Sg приложение={s.sg_kpa} источник={d.get(3) if d else None}")
print(f"5) СП20 Прил. К: сверено {n5}, расхождений {bad5}; со значением Sg: {st.sg_kpa.notna().sum()}")

# 6 -----------------------------------------------------------------------
# Табл. 13.1, 13.2, 15.1, А.1 — против однолистовой формы (её расчет для координаты в строке 31)
sys.path.insert(0, str(ROOT))
import core  # noqa: E402

coord = str(mn[31].get(1)).split("координаты")[-1].strip()
flat, flon = b.parse_coords(coord)
bad6 = n6 = 0
t1, t2 = core.REF["t13_1"], core.REF["t13_2"]
for i in range(12):
    exp = b.num(mn[40].get(3 + i))
    got = core._interp_lat(t1["lats"], t1["rows"][i], flat)
    n6 += 1
    if exp is None or got is None or abs(exp - got) > 0.06:
        bad6 += 1
        problems.append(f"[6] 13.1 {core.MONTHS[i]}: приложение={got} форма={exp}")
for row_, orient in ((33, "Ю"), (34, "ЮВ/ЮЗ"), (35, "В/З"), (36, "СВ/СЗ"), (37, "С")):
    k = t2["orients"].index(orient)
    for i in range(12):
        exp = b.num(mn[row_].get(3 + i))
        got = core._interp_lat(t2["lats"], t2["blocks"][i][k], flat)
        n6 += 1
        if got is None and (exp is None or exp == 0):
            continue  # в таблице СП нет значения на этой широте (форма показывает 0)
        if exp is None or got is None or abs(exp - got) > 0.06:
            bad6 += 1
            problems.append(f"[6] 13.2 {orient} {core.MONTHS[i]}: приложение={got} форма={exp}")
sun = core.sun_table(flat)
dmax = hmax = lmax = 0.0
for i in range(12):  # форма считает склонение приближенно — выводим отклонение для сведения
    dmax = max(dmax, abs(b.num(mn[77].get(7 + i)) - sun[i][0]))
    hmax = max(hmax, abs(b.num(mn[78].get(7 + i)) - sun[i][1]))
    exp = str(mn[81].get(7 + i))  # «8ч 0м»
    h, m_ = [int(x) for x in exp.replace("ч", " ").replace("м", " ").split()]
    lmax = max(lmax, abs(h * 60 + m_ - sun[i][2] * 60))
print(f"   15.1: макс. отклонение от формы — склонение {dmax:.2f}°, высота {hmax:.2f}°, долгота дня {lmax:.0f} мин "
      "(в приложении склонение по алгоритму NOAA)")
form_sub = str(mn[106].get(8)).strip()
sb = st[st["name"] == name].iloc[0]
app_sub = core.climate_subregions(core.rec(sb))
n6 += 1
if app_sub != [form_sub]:
    bad6 += 1
    problems.append(f"[6] А.1 {name}: приложение={app_sub} форма={form_sub}")
print(f"6) Табл. 13.1, 13.2, 15.1, А.1 против формы ({flat}): сверено {n6}, расхождений {bad6}")

print("\nЗамечания:" if problems else "\nЗамечаний нет.")
for p in problems:
    print(" ", p)
