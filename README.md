# Modulsteckbrief-Explorer

Streamlit-Anwendung zur **Auswahl, Abfrage und Ausgabe** der Modulsteckbriefe
einer Remanufacturing-Learning-Factory. Entstanden im Rahmen einer Masterarbeit
(M.Sc.) zum didaktischen Framework für den Kompetenzaufbau im Remanufacturing.

## Funktion

- **Zielgruppen-Filter** über ein kontrolliertes Vokabular (Dropdown, normiert
  aus dem Blatt `LISTEN`), wahlweise mit **UND-/ODER-Verknüpfung**.
- **Freitextsuche** über alle Felder der Steckbriefe.
- **Vollständige Ausgabe** des Steckbriefs (alle 8 Abschnitte) je Modul.
- **CSV-Export** der gefilterten Übersicht.

## Projektstruktur

```
.
├── app.py                          # Streamlit-App
├── requirements.txt
├── data/
│   ├── Arbeitsergebnis_01_Modulsteckbriefe.xlsx   # Quelldaten
│   └── modules.json                # generierte, bereinigte Datenbasis
└── scripts/
    └── convert_xlsx_to_json.py     # Konverter xlsx -> json
```

## Schnellstart (lokal)

```bash
python -m venv .venv && source .venv/bin/activate   # optional
pip install -r requirements.txt

# 1) Datenbasis aus der Excel erzeugen (einmalig bzw. bei Datenänderung)
python scripts/convert_xlsx_to_json.py \
    --input data/Arbeitsergebnis_01_Modulsteckbriefe.xlsx \
    --output data/modules.json

# 2) App starten
streamlit run app.py
```

Die App ist anschließend unter http://localhost:8501 erreichbar.

## Deployment (Streamlit Community Cloud)

1. Repository auf GitHub pushen.
2. Auf [share.streamlit.io](https://share.streamlit.io) das Repo verbinden,
   als Main file `app.py` wählen.
3. `data/modules.json` muss im Repo liegen (wird mit eingecheckt) – die App
   benötigt zur Laufzeit kein Excel.

## Datenmodell

`scripts/convert_xlsx_to_json.py` liest jedes Modul-Blatt (Schema: zwei Spalten,
Feld/Wert) und schreibt eine strukturierte `modules.json`. Die Zielgruppen aus
dem Steckbrief-Freitext werden zusätzlich auf das kontrollierte Vokabular
abgebildet (`zielgruppe_kanon`). Das Mapping ist im Skript transparent
dokumentiert; der Originaltext bleibt erhalten (`zielgruppe_text`).

> Hinweis zum Mapping: „Industrie“ wird mangels eigener LISTEN-Kategorie auf
> **Unternehmer:innen** abgebildet. Diese Heuristik ist in der Arbeit zu
> benennen und bei Bedarf im Konverter (`ZIELGRUPPEN_KANON`) anpassbar.
