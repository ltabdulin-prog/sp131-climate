"""Строительная климатология — СП 131.13330.2025. Streamlit-приложение."""
import pandas as pd
import pydeck as pdk
import streamlit as st

import core

st.set_page_config(page_title="Климатология СП 131.13330.2025", layout="wide")

st.markdown(
    """
    <style>
      #MainMenu, footer, [data-testid="stDecoration"] {visibility: hidden; height: 0;}
      .block-container {padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1500px;}
      h1 {font-size: 1.45rem !important; font-weight: 600 !important; margin-bottom: 0 !important;}
      h2 {font-size: 1.15rem !important; font-weight: 600 !important;}
      h3 {font-size: 1.0rem !important; font-weight: 600 !important; margin-top: 1.1rem !important;}
      .src {color: #666; font-size: 0.82rem; margin-bottom: 0.8rem;}
      .note {color: #555; font-size: 0.8rem; margin-top: -0.4rem;}
      [data-testid="stTable"] table {font-size: 0.84rem;}
      [data-testid="stTable"] th {background: #f3f4f6; font-weight: 600;}
      [data-testid="stTable"] td, [data-testid="stTable"] th {padding: 0.25rem 0.5rem !important;}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load():
    return core.load_data()


stations, prev, osr = load()
prev_by_id = {sid: {r["edition"]: core.rec(r) for r in g.to_dict("records")} for sid, g in prev.groupby("id")}

# ------------------------------------------------------------------ шапка
st.title("Строительная климатология — СП 131.13330.2025")
st.markdown(
    f'<div class="src">Станций: {len(stations)}. Основной источник — СП 131.13330.2025; '
    "прежние редакции (СП 131.13330.2020, 2018, СНиП 23-01-99*) — справочно.</div>",
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------ фильтры
with st.sidebar:
    st.header("Фильтры")
    q = st.text_input("Поиск по названию / адресу / субъекту", placeholder="например, Барнаул")
    regions = sorted(stations["region"].dropna().unique())
    reg = st.multiselect("Субъект РФ", regions, placeholder="все")
    t5 = stations["c_5d092"].dropna()
    t5_rng = st.slider("t наиболее холодной пятидневки 0,92, °С",
                       float(t5.min()), float(t5.max()), (float(t5.min()), float(t5.max())), step=1.0)
    subs = sorted({s.strip() for v in stations["climate_sub"].dropna() for s in v.split(",")},
                  key=lambda x: (len(x.rstrip("АБВГД")), x))
    sub_sel = st.multiselect("Климатический подрайон (условно, табл. А.1)", subs, placeholder="все")
    seis = st.selectbox("Сейсмичность ОСР-2016, карта A (у станции)",
                        ["любая", "≥ 6 баллов", "≥ 7 баллов", "≥ 8 баллов", "≥ 9 баллов"], placeholder="любая")
    st.divider()
    st.header("Точка площадки")
    pt_text = st.text_input("Координаты (широта, долгота)", placeholder="55.7558, 37.6176",
                            help="Таблица сортируется по удаленности от точки; в карточке — сейсмичность для точки.")
    point = core.parse_point(pt_text)
    if pt_text and point is None:
        st.caption("Не удалось распознать координаты. Формат: 55.7558, 37.6176")
    st.divider()
    with_prev = st.checkbox("Показывать прежние редакции (справочно)", value=True)

df = stations.copy()
if q:
    for word in q.lower().replace("ё", "е").split():
        df = df[df["search"].str.contains(word, regex=False)]
if reg:
    df = df[df["region"].isin(reg)]
if t5_rng != (float(t5.min()), float(t5.max())):
    df = df[df["c_5d092"].between(*t5_rng)]
if sub_sel:
    df = df[df["climate_sub"].fillna("").apply(lambda v: any(s in [x.strip() for x in v.split(",")] for s in sub_sel))]
if seis != "любая":
    lim = int(seis.split()[1])
    df = df[pd.to_numeric(df["osr_A"].astype(str).str.extract(r"(\d+)")[0], errors="coerce") >= lim]
if point:
    df = df.assign(dist_km=core.haversine_km(point[0], point[1], df["lat"].values, df["lon"].values).round(1))
    df = df.sort_values("dist_km")

# ------------------------------------------------------------------ выбранная станция
if "sel_id" not in st.session_state:
    st.session_state.sel_id = int(df["id"].iloc[0]) if len(df) else int(stations["id"].iloc[0])
if point and st.session_state.get("_last_point") != point and len(df):
    st.session_state.sel_id = int(df["id"].iloc[0])  # новая точка -> ближайшая станция
if len(df) and st.session_state.sel_id not in set(df["id"]):
    st.session_state.sel_id = int(df["id"].iloc[0])  # выбранная выпала из фильтра -> первая найденная
st.session_state._last_point = point

LIST_COLS = {
    "id": "№", "name": "Пункт", "region": "Субъект",
    "c_5d092": "t 5-дн. 0,92", "c_5d098": "t 5-дн. 0,98", "c_t094": "t 0,94", "c_absmin": "t абс. мин",
    "c_n8": "≤8 °С, сут", "c_t8": "≤8 °С, t ср", "w_t095": "t тепл. 0,95", "w_t098": "t тепл. 0,98",
    "t_year": "t год", "mt": "Mt", "climate_sub": "Подрайон (усл.)", "sg_kpa": "Sg, кПа",
    "osr_A": "ОСР A", "osr_B": "ОСР B", "osr_C": "ОСР C",
}
if point:
    LIST_COLS = {"dist_km": "Расст., км", **LIST_COLS}

tab_list, tab_map = st.tabs(["Таблица станций", "Карта"])

with tab_list:
    view = df[list(LIST_COLS)].rename(columns=LIST_COLS)
    ev = st.dataframe(
        view, hide_index=True, use_container_width=True, height=min(360, 38 + 35 * max(len(view), 1)),
        on_select="rerun", selection_mode="single-row", key="tbl",
        column_config={"№": st.column_config.NumberColumn(width="small", format="%d")},
    )
    sel_rows = ev.selection.rows if ev is not None else []
    sig = (tuple(sel_rows), len(df), tuple(df["id"].head(3)))
    if sel_rows and st.session_state.get("_tbl_sig") != sig:
        st.session_state.sel_id = int(df["id"].iloc[sel_rows[0]])
    st.session_state._tbl_sig = sig
    c1, c2 = st.columns([1, 5])
    c1.download_button("Список в Excel", core.list_to_excel(view), file_name="stations_sp131_2025.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.caption(f"Показано станций: {len(df)}. Выберите строку, чтобы открыть карточку.")

with tab_map:
    param_opts = {
        "t наиболее холодной пятидневки 0,92": "c_5d092",
        "t абсолютная минимальная": "c_absmin",
        "t обеспеченностью 0,98 (теплый период)": "w_t098",
        "Средняя годовая температура": "t_year",
        "Годовая сумма осадков": "pr_year",
        "Средняя годовая скорость ветра": "v_year",
    }
    pname = st.selectbox("Цвет точек по параметру", list(param_opts), key="map_param")
    pcol = param_opts[pname]
    m = df[["id", "name", "region", "lat", "lon", pcol]].dropna(subset=["lat", "lon"]).copy()
    vals = m[pcol].astype(float)
    lo, hi = (vals.min(), vals.max()) if vals.notna().any() else (0, 1)
    rev = pcol in ("pr_year",)

    C_LO, C_MID, C_HI = (33, 102, 172), (235, 235, 235), (178, 24, 43)  # синий — светлый — красный

    def color(v):
        if pd.isna(v) or hi == lo:
            return [140, 140, 140, 200]
        x = (v - lo) / (hi - lo)
        x = 1 - x if rev else x
        a, b, k = (C_LO, C_MID, x * 2) if x < 0.5 else (C_MID, C_HI, (x - 0.5) * 2)
        return [int(a[i] + (b[i] - a[i]) * k) for i in range(3)] + [230]

    m["color"] = vals.apply(color)
    m["val"] = vals.apply(core.fmt)
    layers = [pdk.Layer("ScatterplotLayer", data=m, id="stations", get_position=["lon", "lat"],
                        get_fill_color="color", get_radius=14000, radius_min_pixels=3, radius_max_pixels=9,
                        stroked=True, get_line_color=[90, 90, 90, 160], line_width_min_pixels=0.5,
                        pickable=True, auto_highlight=True)]
    selrow = stations[stations["id"] == st.session_state.sel_id]
    if len(selrow):
        layers.append(pdk.Layer("ScatterplotLayer", data=selrow[["lon", "lat"]], id="selected",
                                get_position=["lon", "lat"], get_fill_color=[0, 0, 0, 0],
                                get_line_color=[0, 0, 0, 255], stroked=True, filled=False,
                                line_width_min_pixels=2, get_radius=30000, radius_min_pixels=8))
    if point:
        layers.append(pdk.Layer("ScatterplotLayer", data=pd.DataFrame([{"lat": point[0], "lon": point[1]}]),
                                id="point", get_position=["lon", "lat"], get_fill_color=[200, 0, 0, 255],
                                get_radius=8000, radius_min_pixels=5))
    focused = bool(q or reg or point) and len(selrow)
    center = (selrow.iloc[0]["lat"], selrow.iloc[0]["lon"]) if focused else (62, 95)
    deck = pdk.Deck(
        layers=layers,
        initial_view_state=pdk.ViewState(latitude=float(center[0]), longitude=float(center[1]),
                                         zoom=4.5 if focused else 2.2),
        tooltip={"text": "{name} ({region})\n" + pname + ": {val}"},
        map_style="light",
    )
    mev = st.pydeck_chart(deck, on_select="rerun", selection_mode="single-object", key="map",
                          use_container_width=True, height=560)
    objs = []
    try:
        objs = (mev.selection or {}).get("objects", {}).get("stations", [])
    except AttributeError:
        objs = []
    msig = tuple(o.get("id") for o in objs)
    if objs and st.session_state.get("_map_sig") != msig:
        st.session_state.sel_id = int(objs[0]["id"])
        st.session_state._map_sig = msig
        st.rerun()
    st.session_state._map_sig = msig
    st.caption(f"Шкала: {core.fmt(lo)} … {core.fmt(hi)} (синий — меньше, красный — больше). "
               "Щелчок по точке открывает карточку станции.")

# ------------------------------------------------------------------ карточка
st.divider()
ids = list(df["id"])
sel = st.session_state.sel_id
if sel not in ids:
    ids = [sel] + ids
names = stations.set_index("id")[["name", "region"]].to_dict("index")
choice = st.selectbox("Станция", ids, index=ids.index(sel),
                      format_func=lambda i: f"{names[i]['name']} — {names[i]['region']} (№ {i})")
if choice != sel:
    st.session_state.sel_id = sel = choice

r = core.rec(stations[stations["id"] == sel].iloc[0])
prev_rows = prev_by_id.get(sel, {}) if with_prev else {}

c1, c2 = st.columns([2, 1])
with c2:
    soil = st.selectbox("Грунт для расчета промерзания (СП 22.13330, п. 5.5.3)", list(core.SOILS))
    kh_on = st.checkbox("Учесть коэффициент kh (п. 5.5.4, табл. 5.2 СП 22.13330)")
    kh, kh_label = None, ""
    if kh_on:
        kht = core.REF["kh"]
        kinds = [row_[0].replace("   ", " ") for row_ in kht["rows"]]
        kind = st.selectbox("Особенности сооружения", kinds)
        temp = st.selectbox("Расчетная среднесуточная температура в помещении у наружных фундаментов",
                            kht["temps"], index=3)
        kh = kht["rows"][kinds.index(kind)][1 + kht["temps"].index(temp)]
        kh_label = f"{kind.lower()}, {temp}"
        st.caption(f"kh = {core.fmt(kh)}")

osr_point = None
if point:
    o, dist = core.nearest_osr(osr, *point)
    osr_point = (o, dist, point)

sections = core.build_sections(r, prev_rows, with_prev=with_prev, soil=soil, kh=kh, osr_point=osr_point,
                               kh_label=kh_label)
title = f"{r['name']} ({r['region']})"

with c1:
    st.subheader(title)
    st.caption(f"{r['address'] or ''} · {r['lat']:.6f}, {r['lon']:.6f}")
    e1, e2, _ = st.columns([1, 1, 1])
    safe = "".join(ch if ch.isalnum() else "_" for ch in r["name"])
    e1.download_button("Экспорт в Excel", core.to_excel(sections, title), file_name=f"climate_{safe}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    e2.download_button("Экспорт в Word", core.to_word(sections, title), file_name=f"climate_{safe}.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)


def show(sec: core.Section):
    st.markdown(f"### {sec.title}")
    d = sec.df.copy()
    for c in d.columns:
        d[c] = d[c].map(core.fmt)
    styler = d.style.hide(axis="index")
    if sec.ref:
        styler = styler.set_properties(subset=[c for c in sec.ref if c in d.columns], color="#7a7a7a")
    st.table(styler)
    if sec.note:
        st.markdown(f'<div class="note">{sec.note}</div>', unsafe_allow_html=True)


def group_of(title: str) -> int:
    if title.startswith("Таблица А") or title.startswith("Таблица Б"):
        return 2
    if title.startswith("Таблица 1") and title[8:10] in ("13", "14", "15", "16"):
        return 1
    if title.startswith("Общие") or title.startswith("Таблица"):
        return 0
    return 3


groups = ["Таблицы 5.1–12.1", "Таблицы 13.1–16.2", "Приложения А, Б", "Расчеты и справочно"]
tabs = st.tabs(groups)
for gi, tab in enumerate(tabs):
    with tab:
        for s in sections:
            if group_of(s.title) == gi:
                show(s)

st.divider()
st.caption(core.SOURCES + " Значения прежних редакций показаны серым и носят справочный характер. "
           "Подрайон по табл. А.1 и производные параметры — расчет приложения, требуют инженерной проверки.")
