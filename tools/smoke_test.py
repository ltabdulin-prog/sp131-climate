"""Прогон карточек всех станций и экспорта без интерфейса: python tools/smoke_test.py"""
import collections
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import core  # noqa: E402

st, prev, osr = core.load_data()
pb = {sid: {r["edition"]: core.rec(r) for r in g.to_dict("records")} for sid, g in prev.groupby("id")}
t = time.time()
err = 0
for _, row in st.iterrows():
    r = core.rec(row)
    try:
        S = core.build_sections(r, pb.get(r["id"], {}), True, "суглинки и глины", 0.6)
        if r["id"] % 50 == 0:
            core.to_excel(S, r["name"])
            core.to_word(S, r["name"])
    except Exception as e:  # noqa: BLE001
        err += 1
        print("ERR", r["name"], repr(e))
print(f"Карточки: {len(st)}, ошибок: {err}, {time.time() - t:.1f} с")
c = collections.Counter(s for v in st.climate_sub.dropna() for s in v.split(", "))
print("Подрайоны (усл.):", dict(sorted(c.items())), "| не определено:", int(st.climate_sub.isna().sum()))

name = sys.argv[1] if len(sys.argv) > 1 else "Барнаул"
b = core.rec(st[st.name == name].iloc[0])
S = core.build_sections(b, pb.get(b["id"], {}), True, "пески гравелистые, крупные и средней крупности", 0.6)
for s in S[-3:]:
    print("\n" + s.title)
    print(s.df.to_string(index=False))
if len(sys.argv) > 2:
    out = Path(sys.argv[2])
    (out / f"{name}.xlsx").write_bytes(core.to_excel(S, name))
    (out / f"{name}.docx").write_bytes(core.to_word(S, name))
    print("Экспорт сохранён в", out)
