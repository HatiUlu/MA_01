# Modulsteckbrief-Explorer

Streamlit-Anwendung zur **Auswahl, Abfrage, Ausgabe und Erfassung** der
Modulsteckbriefe einer Remanufacturing-Learning-Factory. Entstanden im Rahmen
einer Masterarbeit (M.Sc.) zum didaktischen Framework für den Kompetenzaufbau
im Remanufacturing.

## Funktion (drei Tabs)

1. **Filtern & Anzeigen** — Filter über alle ~33 Dimensionen des kontrollierten
   Vokabulars (Blatt `LISTEN`), zusätzlich nach `Level` (aus der Erhebung
   Schwarzer) sowie Freitextsuche. Dimensionen UND-/ODER-verknüpfbar.
   Ausgabe als Übersichtstabelle (CSV-Export) und als vollständige Steckbriefe
   (alle 8 Abschnitte).
2. **Modul erfassen** — leeres Formular mit Dropdowns aus den LISTEN-Werten.
   Neue Module (z. B. das Modul von Prof. Schwarzer) werden **dauerhaft** in
   `data/modules.json` gespeichert.
3. **Verwalten & Export** — manuell erfasste Module einzeln **löschen**;
   komplette `modules.json` als Download.

## Persistenz – wichtiger Hinweis

Das Speichern/Löschen schreibt direkt in `data/modules.json`. Das funktioniert
**lokal** (Ausführung via `streamlit run app.py`) dauerhaft. Auf der
**Streamlit Community Cloud** ist das Dateisystem flüchtig — dort dort gehen
Schreibänderungen beim nächsten Reboot verloren. Für die Cloud daher: Module
lokal erfassen, im Tab „Verwalten & Export" die `modules.json` herunterladen
und ins GitHub-Repo legen.

## Projektstruktur

```
.
├── app.py
├── requirements.txt
├── data/
│   ├── Arbeitsergebnis_01_Modulsteckbriefe.xlsx
│   ├── Arbeitsergebnis_02_Ergaenzung_Schwarzer.xlsx
│   ├── listen.json          # kontrolliertes Vokabular (33 Dimensionen)
│   └── modules.json         # generierte Datenbasis (wird von der App beschrieben)
└── scripts/
    └── convert_xlsx_to_json.py
```

## Schnellstart (lokal – empfohlen für Speichern/Löschen)

```bash
pip install -r requirements.txt

# Datenbasis aus beiden Excel-Dateien erzeugen (einmalig bzw. bei Datenänderung)
python scripts/convert_xlsx_to_json.py \
    --input data/Arbeitsergebnis_01_Modulsteckbriefe.xlsx \
    --schwarzer data/Arbeitsergebnis_02_Ergaenzung_Schwarzer.xlsx \
    --output data/modules.json

streamlit run app.py
```

App unter http://localhost:8501. Beenden mit Strg+C.

## Datenmodell

`scripts/convert_xlsx_to_json.py` liest die 25 Modul-Blätter der Hauptmappe und
reichert sie aus der Schwarzer-Erhebung (`Learning Factory Module Overview`) um
`Level`, `Zustand` und `Ansprechpartner` an. Das kontrollierte Vokabular wird
aus `data/listen.json` (Blatt `LISTEN`) übernommen. Zielgruppen-Freitext wird
zusätzlich auf das Kanon-Vokabular abgebildet (`zielgruppe_kanon`).

> Hinweis: Das Modul von Prof. Schwarzer wird bewusst **nicht** automatisch
> eingetragen — es ist über das Formular im Tab „Modul erfassen" zu erfassen.
