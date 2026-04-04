import pandas as pd
from datetime import datetime

URL = "https://statistiques.public.lu/dam-assets/fr/donnees-autres-formats/indicateurs-court-terme/economie-totale-prix/E5010.xls"

print("📥 Lade STATEC Excel herunter...")

df = pd.read_excel(URL, sheet_name="FR_verso", header=None, engine="xlrd")

# === C1 / C2 (letzte 24 Monate) ===
start_c = None
for i in range(len(df)):
    row = df.iloc[i]
    row_str = " ".join(str(cell).lower() for cell in row if pd.notna(cell))
    if "raccordé" in row_str:
        start_c = i + 3
        print(f"✅ C1/C2 gefunden bei Zeile {i}")
        break

series = []
for i in range(start_c, start_c + 60):
    if i >= len(df): break
    row = df.iloc[i]
    if isinstance(row[0], str) and "/" in str(row[0]) and len(str(row[0])) <= 7:
        date_str = str(row[0]).strip()
        c1 = float(row[3]) if pd.notna(row[3]) else None
        c2 = float(row[5]) if pd.notna(row[5]) else None
        if c1 and c2:
            series.append({"date": date_str, "c1": round(c1, 2), "c2": round(c2, 2)})

data = {}
data["series"] = series[-24:]
latest = series[-1]
data["latest_date"] = latest["date"]
data["c1"] = latest["c1"]
data["c2"] = latest["c2"]

# === COTES + HISTORISCHE TRANCHEN ===
cote_start = None
for i in range(len(df)):
    row = df.iloc[i]
    row_str = " ".join(str(cell).lower() for cell in row if pd.notna(cell))
    if "cotes" in row_str or "d1" in row_str or "échelle mobile" in row_str:
        cote_start = i + 3
        print(f"✅ Cotes gefunden bei Zeile {i}")
        break

tranches = []
for i in range(cote_start, cote_start + 80):
    if i >= len(df): break
    row = df.iloc[i]
    if pd.notna(row[0]) and isinstance(row[0], (int, float)):
        cote_app = float(row[1]) if pd.notna(row[1]) else None
        date_app = None
        if len(row) > 3 and pd.notna(row[3]) and isinstance(row[3], str) and "/" in str(row[3]):
            date_app = str(row[3]).strip()
        if cote_app:
            tranches.append({"date": date_app or "unbekannt", "cote": round(cote_app, 2)})

data["tranches"] = tranches[-30:]   # letzte 30 Tranchen (reicht für den Chart)

# Nächster Schwellenwert & fehlende Punkte
if tranches:
    last = tranches[-1]
    data["last_cote_app"] = last["cote"]
    data["last_cote_ech"] = round(last["cote"] / 1.025, 2)   # zurückgerechnet
    data["next_threshold"] = round(last["cote"] * 1.025, 2)
    data["missing_points"] = round(data["next_threshold"] - data["c2"], 2)
    data["percent_to_next"] = round((data["c2"] / last["cote"] - 1) * 100, 2)

data["updated"] = datetime.utcnow().strftime("%d.%m.%Y %H:%M UTC")

import json
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"✅ Erfolgreich! C2 = {data['c2']} | Nächste Tranche in {data['missing_points']} Punkten")
