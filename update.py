import pandas as pd
import requests
from datetime import datetime

URL = "https://statistiques.public.lu/dam-assets/fr/donnees-autres-formats/indicateurs-court-terme/economie-totale-prix/E5010.xls"

print("📥 Lade STATEC Excel herunter...")

# Excel herunterladen und parsen
df = pd.read_excel(URL, sheet_name="FR_verso", header=None, engine="xlrd")

# =============================================
# ROBUSTE SUCHE NACH C1 / C2 (Index + Moyenne)
# =============================================
start_c = None
for i in range(len(df)):
    row = df.iloc[i]
    if isinstance(row[0], str) and "raccordé" in str(row[0]).lower():
        start_c = i + 2
        print(f"✅ C1-Abschnitt gefunden bei Zeile {i}")
        break

if start_c is None:
    raise ValueError("❌ C1-Abschnitt nicht gefunden. Excel-Struktur hat sich stark geändert.")

# Letzte 24 Monate auslesen
series = []
for i in range(start_c, start_c + 60):
    if i >= len(df):
        break
    row = df.iloc[i]
    if isinstance(row[0], str) and "/" in str(row[0]) and len(str(row[0])) == 7:
        date_str = str(row[0])
        c1 = float(row[3]) if pd.notna(row[3]) else None
        c2 = float(row[5]) if pd.notna(row[5]) else None
        if c1 is not None and c2 is not None:
            series.append({"date": date_str, "c1": round(c1, 2), "c2": round(c2, 2)})

data = {}
data["series"] = series[-24:]          # letzte 24 Monate
latest = series[-1]
data["latest_date"] = latest["date"]
data["c1"] = latest["c1"]
data["c2"] = latest["c2"]

# =============================================
# ROBUSTE SUCHE NACH COTES D'APPLICATION
# =============================================
cote_start = None
for i in range(len(df)):
    row = df.iloc[i]
    if isinstance(row[0], str) and ("cotes d'application" in str(row[0]).lower() or "d1" in str(row[0]).lower()):
        cote_start = i + 3
        print(f"✅ Cotes-Abschnitt gefunden bei Zeile {i}")
        break

if cote_start is None:
    raise ValueError("❌ Cotes-Abschnitt nicht gefunden.")

cotes = []
for i in range(cote_start, cote_start + 50):
    if i >= len(df):
        break
    row = df.iloc[i]
    if pd.notna(row[0]) and isinstance(row[0], (int, float)):
        cote_ech = float(row[0])
        cote_app = float(row[1]) if pd.notna(row[1]) else None
        if cote_app:
            cotes.append({"cote_ech": cote_ech, "cote_app": cote_app})

if cotes:
    last = cotes[-1]
    data["last_cote_app"] = round(last["cote_app"], 2)
    data["last_cote_ech"] = round(last["cote_ech"], 2)
    data["next_threshold"] = round(last["cote_ech"] * 1.025, 2)
    data["percent_to_next"] = round((data["c2"] / last["cote_ech"] - 1) * 100, 2)

# Timestamp
data["updated"] = datetime.utcnow().strftime("%d.%m.%Y %H:%M UTC")

# Speichern
import json
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"✅ Update erfolgreich! Moyenne semestrielle = {data['c2']} (Stand {data['latest_date']})")
