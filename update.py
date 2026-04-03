import pandas as pd
import requests
from datetime import datetime

URL = "https://statistiques.public.lu/dam-assets/fr/donnees-autres-formats/indicateurs-court-terme/economie-totale-prix/E5010.xls"

# Excel herunterladen und parsen (.xls → xlrd)
df = pd.read_excel(URL, sheet_name="FR_verso", header=None, engine="xlrd")

# Suche nach den relevanten Sektionen (Struktur ist seit Jahren stabil)
data = {}

# --- C1 & C2 finden (letzte 24 Monate) ---
for i, row in df.iterrows():
    if isinstance(row[0], str) and "C1. Indice général raccordé" in str(row[0]):
        start_c = i + 2
        break

series = []
for i in range(start_c, start_c + 60):  # max 5 Jahre
    if i >= len(df):
        break
    row = df.iloc[i]
    if isinstance(row[0], str) and "/" in str(row[0]) and len(str(row[0])) == 7:  # z.B. "2026/02"
        date = str(row[0])
        c1 = float(row[3]) if pd.notna(row[3]) else None   # Spalte C1 meist Index 3-4
        c2 = float(row[5]) if pd.notna(row[5]) else None   # C2 meist weiter rechts
        if c1 and c2:
            series.append({"date": date, "c1": round(c1, 2), "c2": round(c2, 2)})

data["series"] = series[-24:]  # letzte 24 Monate

# Aktuellste Werte
latest = series[-1]
data["latest_date"] = latest["date"]
data["c1"] = latest["c1"]
data["c2"] = latest["c2"]

# --- Section D: Cotes d'application & échéance ---
for i, row in df.iterrows():
    if isinstance(row[0], str) and "D1" in str(row[0]) and "Cotes d" in str(row[0]):
        cote_start = i + 3
        break

cotes = []
for i in range(cote_start, cote_start + 50):
    if i >= len(df):
        break
    row = df.iloc[i]
    if pd.notna(row[0]) and isinstance(row[0], (int, float)):
        cote_ech = float(row[0])
        cote_app = float(row[1]) if pd.notna(row[1]) else None
        date_app = str(row[3]) if pd.notna(row[3]) else None
        if cote_app:
            cotes.append({"cote_ech": cote_ech, "cote_app": cote_app, "date_app": date_app})

if cotes:
    last = cotes[-1]
    data["last_cote_app"] = round(last["cote_app"], 2)
    data["last_cote_ech"] = round(last["cote_ech"], 2)
    data["next_threshold"] = round(last["cote_ech"] * 1.025, 2)
    data["percent_to_next"] = round((data["c2"] / last["cote_ech"] - 1) * 100, 2)

# Timestamp
data["updated"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")

# Speichern
import json
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("✅ Update erfolgreich – neueste moyenne semestrielle:", data["c2"])
