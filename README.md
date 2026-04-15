# Labor-Anforderungsanalyse — Albertinen-Krankenhaus

Streamlit-Dashboard zur Analyse medizinischer Labor-Anforderungen nach Einsender (Station/Ambulanz).

## Funktionsweise

Die App liest eine Excel-Datei mit Labor-Anforderungsdaten ein und bietet folgende Ansichten:

| Sektion | Inhalt |
|---|---|
| **Überblick** | KPIs + Volumen aller Einsender im Balkendiagramm |
| **Top 30 Anforderungen je Einsender** | Welche Untersuchungen ordert eine Station am häufigsten? |
| **Top 30 Einsender je Anforderung** | Welche Stationen fordern eine bestimmte Untersuchung an? |
| **Mekko-Chart** | Flächendiagramm: Einsender-Volumen × Anforderungsmix |
| **Heatmap** | Kreuzmatrix Einsender × Anforderungen |
| **Pareto-Analyse** | Kumulative 80/20-Verteilung nach Einsender oder Anforderung |
| **Vergleich zweier Einsender** | Side-by-side Top-15 Anforderungen |
| **Daten-Explorer** | Freitextsuche + CSV-Export |

Überall umschaltbar zwischen **Punktsumme** und **Anzahl**. Einträge mit Punktsumme = 0 (interne Steuerkennzeichen) werden standardmäßig ausgefiltert.

## Input-Format

Im Ordner `input/` liegt **eine Excel-Datei pro Labor**. Der Dateiname (ohne `.xlsx`) erscheint als Laborname in der Seitenleiste, z. B.:

```
input/
  Albertinen-KH.xlsx
  Fenner Labor.xlsx
  MVZ Hamburg.xlsx
```

Jede Datei muss **mindestens diese 4 Spalten** enthalten (exakt so benannt, weitere Spalten werden ignoriert):

| Spaltenname | Typ | Beschreibung |
|---|---|---|
| `Einsender` | Text | Name der einsendenden Station / Ambulanz |
| `Analyse/Leistung` | Text | Name der Laboruntersuchung |
| `Punktsumme` | Zahl | Summe der GOÄ-Punkte (0 = kein Abrechnungswert) |
| `Anzahl` | Zahl | Anzahl der durchgeführten Untersuchungen |

Fehlen eine oder mehrere dieser Spalten, zeigt die App eine Fehlermeldung. Jede Zeile entspricht einer eindeutigen Kombination aus Einsender und Analyse/Leistung.

## Installation & lokaler Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Passwort wird aus der Umgebungsvariable `APP_PASSWORD` gelesen. Für lokale Entwicklung eine `.env`-Datei anlegen (siehe `.env.example`).

## Deployment auf Railway

1. Repo auf GitHub pushen
2. Railway: **New Project → Deploy from GitHub**
3. Umgebungsvariable setzen: `APP_PASSWORD = <passwort>`
4. Railway erkennt das `Procfile` automatisch
