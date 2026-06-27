"""
Modulsteckbrief-Explorer
========================
Streamlit-Anwendung zur Auswahl, Abfrage und Ausgabe der Modulsteckbriefe
einer Remanufacturing-Learning-Factory.

Datenbasis: data/modules.json (erzeugt aus der Excel-Arbeitsmappe via
scripts/convert_xlsx_to_json.py).

Filterlogik:
  - Zielgruppe als kontrolliertes Vokabular (Dropdown, normiert) UND/ODER
  - Freitextsuche über alle Steckbrief-Felder.
Ausgabe: vollständiger Steckbrief (alle 8 Abschnitte) je Modul.
"""

from pathlib import Path
import json

import pandas as pd
import streamlit as st

DATA_PATH = Path(__file__).parent / "data" / "modules.json"

st.set_page_config(
    page_title="Modulsteckbrief-Explorer",
    page_icon="🏭",
    layout="wide",
)

# --- Styling (dezent, fachlich) ----------------------------------------------
st.markdown(
    """
    <style>
      .stApp { background: #f7f7f5; }
      h1, h2, h3 { color: #1c2b36; }
      .modul-karte {
          background: #ffffff; border: 1px solid #e3e3df;
          border-radius: 10px; padding: 1.2rem 1.4rem; margin-bottom: 1rem;
      }
      .badge {
          display:inline-block; background:#e8eef2; color:#1c4966;
          border-radius: 999px; padding: 2px 10px; margin: 2px 4px 2px 0;
          font-size: 0.78rem;
      }
      .sektion-titel {
          color:#0d6b5e; font-weight:600; margin-top:0.9rem;
          border-bottom:1px solid #ececec; padding-bottom:2px;
      }
      .feld-label { color:#5a6b75; font-size:0.82rem; }
      .feld-wert  { color:#1c2b36; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def lade_daten():
    with open(DATA_PATH, encoding="utf-8") as f:
        return json.load(f)


def modul_passt_freitext(modul: dict, suchbegriff: str) -> bool:
    """Sucht den Begriff (case-insensitiv) über alle Steckbrief-Felder."""
    if not suchbegriff:
        return True
    nadel = suchbegriff.lower()
    heuhaufen = " ".join(str(v) for v in modul.get("felder", {}).values()).lower()
    heuhaufen += " " + modul.get("blatt", "").lower()
    return nadel in heuhaufen


def render_steckbrief(modul: dict):
    """Gibt den vollständigen Steckbrief (alle Abschnitte) aus."""
    st.markdown(f"### {modul['name'] or modul['blatt']}")

    # Kopf: ID + normierte Zielgruppen als Badges
    kopf = []
    if modul.get("modul_id"):
        kopf.append(f"**ID:** {modul['modul_id']}")
    if modul.get("version"):
        kopf.append(f"**Stand:** {modul['version']}")
    if kopf:
        st.caption("  ·  ".join(kopf))

    if modul.get("zielgruppe_kanon"):
        badges = "".join(
            f"<span class='badge'>{z}</span>" for z in modul["zielgruppe_kanon"]
        )
        st.markdown(
            f"<div class='feld-label'>Zielgruppen (normiert)</div>{badges}",
            unsafe_allow_html=True,
        )

    # Alle Abschnitte in Originalreihenfolge
    for sektion in data["meta"]["abschnitte"]:
        paare = modul.get("abschnitte", {}).get(sektion)
        if not paare:
            continue
        st.markdown(
            f"<div class='sektion-titel'>{sektion}</div>", unsafe_allow_html=True
        )
        for paar in paare:
            st.markdown(
                f"<div class='feld-label'>{paar['label']}</div>"
                f"<div class='feld-wert'>{paar['wert']}</div>",
                unsafe_allow_html=True,
            )


# --- Daten laden -------------------------------------------------------------
try:
    data = lade_daten()
except FileNotFoundError:
    st.error(
        "Datenbasis nicht gefunden. Bitte zuerst den Konverter ausführen:\n\n"
        "`python scripts/convert_xlsx_to_json.py "
        "--input data/Arbeitsergebnis_01_Modulsteckbriefe.xlsx "
        "--output data/modules.json`"
    )
    st.stop()

module = data["module"]
zg_optionen = data["meta"]["zielgruppen_kanon_vorhanden"]

# --- Kopfbereich -------------------------------------------------------------
st.title("Modulsteckbrief-Explorer")
st.caption(
    f"Remanufacturing-Learning-Factory · {data['meta']['anzahl_module']} Module · "
    f"Quelle: {data['meta']['quelle']}"
)

# --- Sidebar: Filter ---------------------------------------------------------
with st.sidebar:
    st.header("Filter")

    gewaehlte_zg = st.multiselect(
        "Zielgruppe (normiert)",
        options=zg_optionen,
        help="Kontrolliertes Vokabular aus dem Blatt LISTEN. "
        "Mehrfachauswahl möglich.",
    )

    modus = st.radio(
        "Verknüpfung bei mehreren Zielgruppen",
        ["ODER (mindestens eine)", "UND (alle gewählten)"],
        help="ODER: Modul spricht mindestens eine gewählte Zielgruppe an. "
        "UND: Modul deckt alle gewählten Zielgruppen ab.",
    )

    suchbegriff = st.text_input(
        "Freitextsuche",
        placeholder="z. B. YOLO, Demontage, Bloom, KI …",
        help="Durchsucht alle Felder der Steckbriefe.",
    )

    st.divider()
    st.caption(
        "Hinweis: Zielgruppen sind aus dem Steckbrief-Freitext normiert "
        "(„Industrie“ → Unternehmer:innen). Der Originaltext bleibt im "
        "Steckbrief sichtbar."
    )

# --- Filterung ---------------------------------------------------------------
def passt_zielgruppe(modul: dict) -> bool:
    if not gewaehlte_zg:
        return True
    kanon = set(modul.get("zielgruppe_kanon", []))
    if modus.startswith("ODER"):
        return bool(kanon & set(gewaehlte_zg))
    return set(gewaehlte_zg).issubset(kanon)


treffer = [
    m for m in module
    if passt_zielgruppe(m) and modul_passt_freitext(m, suchbegriff)
]

# --- Ergebnisübersicht -------------------------------------------------------
st.subheader(f"{len(treffer)} von {len(module)} Modulen")

if not treffer:
    st.info("Keine Module entsprechen den aktuellen Filtern. Filter lockern.")
    st.stop()

# Übersichtstabelle
uebersicht = pd.DataFrame(
    [
        {
            "Modul": m["name"] or m["blatt"],
            "Zielgruppen": ", ".join(m["zielgruppe_kanon"]),
            "Bloom": m.get("bloom_stufen", ""),
            "Dauer": m.get("dauer", ""),
            "Status": m.get("status", ""),
        }
        for m in treffer
    ]
)
st.dataframe(uebersicht, use_container_width=True, hide_index=True)

# CSV-Export der gefilterten Übersicht
st.download_button(
    "Gefilterte Übersicht als CSV",
    uebersicht.to_csv(index=False).encode("utf-8-sig"),
    file_name="modulauswahl.csv",
    mime="text/csv",
)

st.divider()

# --- Detailausgabe (voller Steckbrief) ---------------------------------------
st.subheader("Vollständige Steckbriefe")

namen = [m["name"] or m["blatt"] for m in treffer]
auswahl = st.selectbox(
    "Modul für die Detailansicht wählen",
    options=["Alle anzeigen"] + namen,
)

if auswahl == "Alle anzeigen":
    for m in treffer:
        with st.container():
            st.markdown("<div class='modul-karte'>", unsafe_allow_html=True)
            render_steckbrief(m)
            st.markdown("</div>", unsafe_allow_html=True)
else:
    m = treffer[namen.index(auswahl)]
    st.markdown("<div class='modul-karte'>", unsafe_allow_html=True)
    render_steckbrief(m)
    st.markdown("</div>", unsafe_allow_html=True)
