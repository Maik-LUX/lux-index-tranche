import pandas as pd
from datetime import datetime

URL = "https://statistiques.public.lu/dam-assets/fr/donnees-autres-formats/indicateurs-court-terme/economie-totale-prix/E5010.xls"

print("📥 Lade aktuelle STATEC E5010.xls herunter...")

df = pd.read_excel(URL, sheet_name="FR_verso", header=None, engine="xlrd")

# =============================================
# SUCHE NACH C1/C2 (sehr robust – sucht im ganzen Row)
# =============================================
start_c = None
for i in range(len(df)):
    row = df.iloc[i]
    row_str = " ".join(str(cell).lower() for cell in row if pd.notna(cell))
    if "raccordé" in row_str:
        start_c = i + 3          # Daten beginnen 3 Zeilen darunter
        print(f"✅ C1/C2-Abschnitt gefunden bei Zeile {i}")
        break

if start_c is None:
    print("❌ C1/C2 nicht gefunden – hier die ersten 50 Zeilen Spalte A zur Diagnose:")
    print(df[0].head(50).to_string())
    raise ValueError("C1-Abschnitt nicht gefunden")

# Daten auslesen (letzte 30 Monate reichen)
series = []
for i in range(start_c, start_c + 60):
    if i >= len(df):
        break
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

# =============================================
# SUCHE NACH COTES D'APPLICATION
# =============================================
cote_start = None
for i in range(len(df)):
    row = df.iloc[i]
    row_str = " ".join(str(cell).lower() for cell in row if pd.notna(cell))
    if "cotes" in row_str or "d1" in row_str or "échelle mobile" in row_str:
        cote_start = i + 3
        print(f"✅ Cotes-Abschnitt gefunden bei Zeile {i}")
        break

if cote_start is None:
    raise ValueError("Cotes-Abschnitt nicht gefunden")

cotes = []
for i in range(cote_start, cote_start + 60):
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

data["updated"] = datetime.utcnow().strftime("%d.%m.%Y %H:%M UTC")

import json
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"✅ ERFOLGREICH! Moyenne semestrielle = {data['c2']} (Stand {data['latest_date']})")
print(f"Nächster Schwellenwert: {data['next_threshold']}")
