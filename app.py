"""
Modulsteckbrief-Explorer
========================
Streamlit-Anwendung zur Auswahl, Abfrage, Ausgabe UND Erfassung der
Modulsteckbriefe einer Remanufacturing-Learning-Factory.

Drei Tabs:
  1. Filtern & Anzeigen  – Filter über alle ~33 LISTEN-Dimensionen + Freitext.
  2. Modul erfassen      – leeres Formular mit allen LISTEN-Dropdowns;
                           speichert dauerhaft in data/modules.json (lokal).
  3. Verwalten & Export  – manuell erfasste Module löschen; JSON-Export.

Persistenz: lokaler Schreibzugriff auf data/modules.json (Weg 3).
Hinweis: Auf Streamlit Community Cloud ist dieser Schreibzugriff NICHT
dauerhaft (flüchtiges Dateisystem). Für die lokale Nutzung via
`streamlit run app.py` funktioniert das Speichern/Löschen wie erwartet.
"""

from pathlib import Path
from datetime import date
import json

import pandas as pd
import streamlit as st

DATA_PATH = Path(__file__).parent / "data" / "modules.json"

st.set_page_config(page_title="Modulsteckbrief-Explorer", page_icon="🏭",
                   layout="wide")

# --- Styling -----------------------------------------------------------------
st.markdown(
    """
    <style>
      .stApp { background: #f7f7f5; }
      h1, h2, h3 { color: #1c2b36; }
      .modul-karte { background:#fff; border:1px solid #e3e3df;
          border-radius:10px; padding:1.2rem 1.4rem; margin-bottom:1rem; }
      .badge { display:inline-block; background:#e8eef2; color:#1c4966;
          border-radius:999px; padding:2px 10px; margin:2px 4px 2px 0;
          font-size:0.78rem; }
      .badge-manuell { background:#e7f3ec; color:#0d6b5e; }
      .sektion-titel { color:#0d6b5e; font-weight:600; margin-top:0.9rem;
          border-bottom:1px solid #ececec; padding-bottom:2px; }
      .feld-label { color:#5a6b75; font-size:0.82rem; }
      .feld-wert  { color:#1c2b36; }
    </style>
    """,
    unsafe_allow_html=True,
)


# --- Daten laden / speichern -------------------------------------------------
def lade_daten() -> dict:
    with open(DATA_PATH, encoding="utf-8") as f:
        return json.load(f)


def speichere_daten(daten: dict):
    daten["meta"]["anzahl_module"] = len(daten["module"])
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(daten, f, ensure_ascii=False, indent=2)


# Im Session-State halten, damit Änderungen sofort sichtbar sind
if "daten" not in st.session_state:
    try:
        st.session_state.daten = lade_daten()
    except FileNotFoundError:
        st.error(
            "Datenbasis nicht gefunden. Bitte zuerst den Konverter ausführen:\n\n"
            "`python scripts/convert_xlsx_to_json.py "
            "--input data/Arbeitsergebnis_01_Modulsteckbriefe.xlsx "
            "--schwarzer data/Arbeitsergebnis_02_Ergaenzung_Schwarzer.xlsx "
            "--output data/modules.json`"
        )
        st.stop()

daten = st.session_state.daten
module = daten["module"]
listen = daten["meta"].get("listen", {})
abschnitte_reihenfolge = daten["meta"]["abschnitte"]

# Zuordnung: LISTEN-Dimension -> Steckbrief-Feld(er), in denen gesucht wird.
# Mehrere Felder möglich; Filter greift als Teilstring (robust ggü. Freitext).
DIM_ZU_FELDER = {
    "Zielgruppe": ["Zielgruppe(n)"],
    "Fachbezogene Lerninhalte": ["Lerninhalte (Stichworte)", "Technologiebezug"],
    "Teilnehmerkonstellation": ["Teilnehmerkonstellation"],
    "Hauptzweck": ["Hauptzweck"],
    "Kompetenzklasse": ["Kompetenzklassen (Erpenbeck)"],
    "Lernszenario": ["Lernszenario-Strategie"],
    "Autonomiegrad": ["Autonomiegrad"],
    "Trainerrolle": ["Trainerrolle"],
    "Lernaktivitätsart": ["Lernaktivität"],
    "Evaluationsebenen": ["Evaluationsebene"],
    "Evaluationsmethoden": ["Prüfungs-/Bewertungsform"],
    "Abgebildete Produktlebenszyklusphase": ["Prozessbezug (Lebenszyklus)"],
    "Prozesstyp": ["Prozesstyp / Variabilität"],
    "Materialität": ["Lernobjekt"],
    "Variantenanzahl": ["Lernobjekt", "Produktkomponenten"],
    "Durchschnittsdauer Lernmodul": ["Dauer"],
    "Teilnehmende pro Lernmodul": ["Teilnehmerzahl"],
}


def feldwerte(modul: dict, felder: list[str]) -> str:
    """Verkettet die Inhalte mehrerer Steckbrief-Felder (kleingeschrieben)."""
    return " ".join(
        str(modul.get("felder", {}).get(f, "")) for f in felder
    ).lower()


def passt_dimension(modul: dict, dim: str, gewaehlt: list[str]) -> bool:
    """Modul passt, wenn EINER der gewählten Werte im zugehörigen Feld steht."""
    if not gewaehlt:
        return True
    felder = DIM_ZU_FELDER.get(dim, [])
    if not felder:
        return True
    hay = feldwerte(modul, felder)
    # zusätzlich manuell gesetzte normierte Werte prüfen
    manuelle = " ".join(modul.get("listen_auswahl", {}).get(dim, [])).lower()
    hay = hay + " " + manuelle
    return any(w.lower() in hay for w in gewaehlt)


def freitext_treffer(modul: dict, q: str) -> bool:
    if not q:
        return True
    hay = " ".join(str(v) for v in modul.get("felder", {}).values()).lower()
    hay += " " + modul.get("blatt", "").lower()
    return q.lower() in hay


def render_steckbrief(modul: dict):
    st.markdown(f"### {modul.get('name') or modul.get('blatt')}")
    kopf = []
    if modul.get("modul_id"):
        kopf.append(f"**ID:** {modul['modul_id']}")
    if modul.get("version"):
        kopf.append(f"**Stand:** {modul['version']}")
    if kopf:
        st.caption("  ·  ".join(kopf))

    # Zusatz-Badges
    zeile = ""
    for z in modul.get("zielgruppe_kanon", []):
        zeile += f"<span class='badge'>{z}</span>"
    if modul.get("level"):
        zeile += f"<span class='badge'>Level: {modul['level']}</span>"
    if modul.get("manuell"):
        zeile += "<span class='badge badge-manuell'>manuell erfasst</span>"
    if zeile:
        st.markdown(zeile, unsafe_allow_html=True)

    if modul.get("zustand"):
        st.caption(f"Zustand: {modul['zustand']}")
    if modul.get("ansprechpartner"):
        st.caption(f"Ansprechpartner: {modul['ansprechpartner']}")

    for sektion in abschnitte_reihenfolge:
        paare = modul.get("abschnitte", {}).get(sektion)
        if not paare:
            continue
        st.markdown(f"<div class='sektion-titel'>{sektion}</div>",
                    unsafe_allow_html=True)
        for paar in paare:
            st.markdown(
                f"<div class='feld-label'>{paar['label']}</div>"
                f"<div class='feld-wert'>{paar['wert']}</div>",
                unsafe_allow_html=True,
            )

    # Bei manuell erfassten Modulen: normierte LISTEN-Auswahl zeigen
    if modul.get("listen_auswahl"):
        st.markdown("<div class='sektion-titel'>Normierte Klassifikation "
                    "(LISTEN)</div>", unsafe_allow_html=True)
        for dim, werte in modul["listen_auswahl"].items():
            if werte:
                st.markdown(
                    f"<div class='feld-label'>{dim}</div>"
                    f"<div class='feld-wert'>{', '.join(werte)}</div>",
                    unsafe_allow_html=True,
                )


# --- Kopf --------------------------------------------------------------------
st.title("Modulsteckbrief-Explorer")
st.caption(
    f"Remanufacturing-Learning-Factory · {len(module)} Module · "
    "Filter über das kontrollierte Vokabular (Blatt LISTEN)"
)

tab_filter, tab_erfassen, tab_verwalten = st.tabs(
    ["🔎 Filtern & Anzeigen", "➕ Modul erfassen", "🗂 Verwalten & Export"]
)

# =============================================================================
# TAB 1 – FILTERN & ANZEIGEN
# =============================================================================
with tab_filter:
    with st.sidebar:
        st.header("Filter")
        suchbegriff = st.text_input(
            "Freitextsuche", placeholder="z. B. YOLO, Demontage, Bloom …"
        )

        # Zusatzdimensionen aus Schwarzer (Level / Zustand)
        level_opt = sorted({m.get("level", "") for m in module if m.get("level")})
        gewaehlt_level = st.multiselect("Level (Schwarzer)", level_opt)

        st.divider()
        st.caption("Filter nach LISTEN-Dimensionen")
        verknuepfung = st.radio(
            "Verknüpfung der Dimensionen",
            ["UND (alle Filter erfüllt)", "ODER (mind. ein Filter)"],
            help="Bezieht sich auf das Zusammenspiel verschiedener Dimensionen.",
        )

        # Alle LISTEN-Dimensionen als (einklappbare) Mehrfachauswahl
        aktive_filter: dict[str, list[str]] = {}
        for dim, werte in listen.items():
            with st.expander(dim, expanded=False):
                wahl = st.multiselect(
                    dim, werte, key=f"filt_{dim}", label_visibility="collapsed"
                )
                if wahl:
                    aktive_filter[dim] = wahl

    # Filterlogik
    def modul_passt(m: dict) -> bool:
        if not freitext_treffer(m, suchbegriff):
            return False
        if gewaehlt_level and m.get("level") not in gewaehlt_level:
            return False
        if not aktive_filter:
            return True
        ergebnisse = [passt_dimension(m, d, w) for d, w in aktive_filter.items()]
        return all(ergebnisse) if verknuepfung.startswith("UND") else any(ergebnisse)

    treffer = [m for m in module if modul_passt(m)]

    st.subheader(f"{len(treffer)} von {len(module)} Modulen")

    if aktive_filter or gewaehlt_level or suchbegriff:
        aktiv = [f"{d}: {', '.join(w)}" for d, w in aktive_filter.items()]
        if gewaehlt_level:
            aktiv.append(f"Level: {', '.join(gewaehlt_level)}")
        if suchbegriff:
            aktiv.append(f"Suche: „{suchbegriff}“")
        st.caption("Aktive Filter — " + "  •  ".join(aktiv))

    if not treffer:
        st.info("Keine Module entsprechen den Filtern. Filter lockern.")
    else:
        uebersicht = pd.DataFrame([
            {
                "Modul": m.get("name") or m.get("blatt"),
                "Zielgruppen": ", ".join(m.get("zielgruppe_kanon", [])),
                "Level": m.get("level", ""),
                "Bloom": m.get("bloom_stufen", ""),
                "Dauer": m.get("dauer", ""),
                "Quelle": "manuell" if m.get("manuell") else "Excel",
            }
            for m in treffer
        ])
        st.dataframe(uebersicht, use_container_width=True, hide_index=True)
        st.download_button(
            "Gefilterte Übersicht als CSV",
            uebersicht.to_csv(index=False).encode("utf-8-sig"),
            file_name="modulauswahl.csv", mime="text/csv",
        )

        st.divider()
        st.subheader("Vollständige Steckbriefe")
        namen = [m.get("name") or m.get("blatt") for m in treffer]
        auswahl = st.selectbox("Modul für die Detailansicht",
                               ["Alle anzeigen"] + namen)
        ziel = treffer if auswahl == "Alle anzeigen" else \
            [treffer[namen.index(auswahl)]]
        for m in ziel:
            st.markdown("<div class='modul-karte'>", unsafe_allow_html=True)
            render_steckbrief(m)
            st.markdown("</div>", unsafe_allow_html=True)

# =============================================================================
# TAB 2 – MODUL ERFASSEN (leeres Formular, alle LISTEN-Dimensionen)
# =============================================================================
with tab_erfassen:
    st.subheader("Neues Modul erfassen")
    st.caption(
        "Leeres Formular. Freitextfelder für die Steckbrief-Abschnitte, "
        "Dropdowns für die normierte LISTEN-Klassifikation. Speichern schreibt "
        "dauerhaft in data/modules.json (lokal)."
    )

    with st.form("modul_erfassen", clear_on_submit=False):
        st.markdown("**1 Identifikation**")
        c1, c2 = st.columns(2)
        f_id = c1.text_input("Modul-ID / Kürzel")
        f_version = c2.text_input("Version / Stand",
                                  value=f"V_01 / {date.today().strftime('%d.%m.%Y')}")
        f_name = st.text_input("Modulname *")
        f_autor = st.text_input("Autor:in / verantwortlich")
        f_ansprech = st.text_input("Ansprechpartner")

        st.markdown("**2 Zielgruppe & Adressierung**")
        f_zielgruppe = st.multiselect("Zielgruppe(n)",
                                      listen.get("Zielgruppe", []))
        f_konstellation = st.selectbox(
            "Teilnehmerkonstellation",
            [""] + listen.get("Teilnehmerkonstellation", []))
        f_hauptzweck = st.multiselect("Hauptzweck", listen.get("Hauptzweck", []))
        f_vorwissen = st.text_input("Vorwissen / Voraussetzung")

        st.markdown("**3 Lernziele & Kompetenzen**")
        f_lernziel = st.text_area("Übergeordnetes Lernziel")
        f_kompetenz = st.multiselect("Kompetenzklassen (Erpenbeck)",
                                     listen.get("Kompetenzklasse", []))
        f_bloom = st.text_input("Kognitive Stufen (Bloom)")
        f_feinlernziele = st.text_area("Feinlernziele")

        st.markdown("**4 Inhalt & Prozessbezug**")
        f_lebenszyklus = st.multiselect(
            "Abgebildete Produktlebenszyklusphase",
            listen.get("Abgebildete Produktlebenszyklusphase", []))
        f_lerninhalte = st.multiselect("Fachbezogene Lerninhalte",
                                       listen.get("Fachbezogene Lerninhalte", []))
        f_technologien = st.multiselect("Integrierte digitale Technologien",
                                        listen.get("Integrierte digitale Technologien", []))
        f_materialitaet = st.selectbox("Materialität",
                                       [""] + listen.get("Materialität", []))

        st.markdown("**5 Didaktik & Methodik**")
        f_szenario = st.selectbox("Lernszenario-Strategie",
                                  [""] + listen.get("Lernszenario", []))
        f_autonomie = st.selectbox("Autonomiegrad",
                                   [""] + listen.get("Autonomiegrad", []))
        f_trainer = st.selectbox("Trainerrolle",
                                 [""] + listen.get("Trainerrolle", []))
        f_aktivitaet = st.multiselect("Lernaktivität",
                                      listen.get("Lernaktivitätsart", []))

        st.markdown("**6 Organisation & Einordnung**")
        c3, c4 = st.columns(2)
        f_dauer = c3.selectbox("Dauer",
                               [""] + listen.get("Durchschnittsdauer Lernmodul", []))
        f_teilnehmer = c4.selectbox("Teilnehmende pro Lernmodul",
                                    [""] + listen.get("Teilnehmende pro Lernmodul", []))
        f_setting = st.selectbox("Lernumgebung",
                                 [""] + listen.get("Lernumgebung", []))
        f_ausstattung = st.text_area("Benötigte Ausstattung")

        st.markdown("**7 Evaluation & Erfolgskriterien**")
        f_evalebene = st.multiselect("Evaluationsebene",
                                     listen.get("Evaluationsebenen", []))
        f_evalmethode = st.multiselect("Evaluationsmethoden",
                                       listen.get("Evaluationsmethoden", []))
        f_erfolg = st.text_input("Erfolgskriterium")

        st.markdown("**8 Wirtschaftlichkeit & Trägerschaft**")
        f_betreiber = st.selectbox("Betreiber",
                                   [""] + listen.get("Betreiber", []))
        f_traegerschaft = st.text_input("Trägerschaft / Partner")
        f_geschaeftsmodell = st.selectbox("Geschäftsmodell (Training)",
                                          [""] + listen.get("Trainingsmodelle", []))
        f_partnerschaften = st.multiselect("Schlüsselpartnerschaften",
                                           listen.get("Schlüsselpartnerschaften", []))
        f_einrichtungskosten = st.selectbox("Einrichtungskosten",
                                            [""] + listen.get("Einrichtungskosten", []))
        f_zustand = st.text_input("Zustand / Reifegrad")

        absenden = st.form_submit_button("Modul speichern", type="primary")

    if absenden:
        if not f_name.strip():
            st.error("Bitte mindestens einen Modulnamen angeben.")
        else:
            # Steckbrief-Struktur aufbauen (nur befüllte Felder)
            def fuelle(paare):
                return [{"label": k, "wert": v} for k, v in paare
                        if v not in (None, "", [])]

            def joinl(x):
                return " | ".join(x) if isinstance(x, list) else x

            abschnitte_neu = {
                "1  Identifikation": fuelle([
                    ("Modul-ID / Kürzel", f_id), ("Modulname", f_name),
                    ("Version / Stand", f_version),
                    ("Autor:in / verantwortlich", f_autor)]),
                "2  Zielgruppe & Adressierung": fuelle([
                    ("Zielgruppe(n)", joinl(f_zielgruppe)),
                    ("Vorwissen / Voraussetzung", f_vorwissen),
                    ("Teilnehmerkonstellation", f_konstellation),
                    ("Hauptzweck", joinl(f_hauptzweck))]),
                "3  Lernziele & Kompetenzen": fuelle([
                    ("Übergeordnetes Lernziel", f_lernziel),
                    ("Kompetenzklassen (Erpenbeck)", joinl(f_kompetenz)),
                    ("Kognitive Stufen (Bloom)", f_bloom),
                    ("Feinlernziele", f_feinlernziele)]),
                "4  Inhalt & Prozessbezug (Remanufacturing)": fuelle([
                    ("Prozessbezug (Lebenszyklus)", joinl(f_lebenszyklus)),
                    ("Lerninhalte (Stichworte)", joinl(f_lerninhalte)),
                    ("Technologiebezug", joinl(f_technologien)),
                    ("Lernobjekt", f_materialitaet)]),
                "5  Didaktik & Methodik": fuelle([
                    ("Lernszenario-Strategie", f_szenario),
                    ("Autonomiegrad", f_autonomie),
                    ("Trainerrolle", f_trainer),
                    ("Lernaktivität", joinl(f_aktivitaet))]),
                "6  Organisation & Einordnung": fuelle([
                    ("Dauer", f_dauer), ("Teilnehmerzahl", f_teilnehmer),
                    ("Ort / Setting", f_setting),
                    ("Benötigte Ausstattung", f_ausstattung)]),
                "7  Evaluation & Erfolgskriterien": fuelle([
                    ("Evaluationsebene", joinl(f_evalebene)),
                    ("Prüfungs-/Bewertungsform", joinl(f_evalmethode)),
                    ("Erfolgskriterium", f_erfolg)]),
                "8  Wirtschaftlichkeit & Trägerschaft": fuelle([
                    ("Trägerschaft / Partner", f_traegerschaft),
                    ("Geschäftsmodell", f_geschaeftsmodell),
                    ("Status / Reifegrad", f_zustand)]),
            }
            felder_flach = {p["label"]: p["wert"]
                            for sek in abschnitte_neu.values() for p in sek}

            neues = {
                "blatt": f_id or f_name,
                "modul_id": f_id, "name": f_name, "version": f_version,
                "autor": f_autor,
                "zielgruppe_text": joinl(f_zielgruppe),
                "zielgruppe_kanon": f_zielgruppe,
                "lernziel": f_lernziel,
                "kompetenzklassen": joinl(f_kompetenz),
                "bloom_stufen": f_bloom, "hauptzweck": joinl(f_hauptzweck),
                "dauer": f_dauer, "status": f_zustand,
                "level": "", "zustand": f_zustand,
                "ansprechpartner": f_ansprech,
                "manuell": True,
                "quelle_datei": "manuelle Eingabe",
                "felder": felder_flach,
                "abschnitte": {k: v for k, v in abschnitte_neu.items() if v},
                # vollständige normierte Mehrfachauswahl für Filter & Anzeige
                "listen_auswahl": {
                    "Zielgruppe": f_zielgruppe,
                    "Hauptzweck": f_hauptzweck,
                    "Kompetenzklasse": f_kompetenz,
                    "Fachbezogene Lerninhalte": f_lerninhalte,
                    "Integrierte digitale Technologien": f_technologien,
                    "Abgebildete Produktlebenszyklusphase": f_lebenszyklus,
                    "Lernaktivitätsart": f_aktivitaet,
                    "Evaluationsebenen": f_evalebene,
                    "Evaluationsmethoden": f_evalmethode,
                    "Schlüsselpartnerschaften": f_partnerschaften,
                    "Teilnehmerkonstellation": [f_konstellation] if f_konstellation else [],
                    "Lernszenario": [f_szenario] if f_szenario else [],
                    "Autonomiegrad": [f_autonomie] if f_autonomie else [],
                    "Trainerrolle": [f_trainer] if f_trainer else [],
                    "Materialität": [f_materialitaet] if f_materialitaet else [],
                    "Lernumgebung": [f_setting] if f_setting else [],
                    "Betreiber": [f_betreiber] if f_betreiber else [],
                    "Trainingsmodelle": [f_geschaeftsmodell] if f_geschaeftsmodell else [],
                    "Durchschnittsdauer Lernmodul": [f_dauer] if f_dauer else [],
                    "Teilnehmende pro Lernmodul": [f_teilnehmer] if f_teilnehmer else [],
                    "Einrichtungskosten": [f_einrichtungskosten] if f_einrichtungskosten else [],
                },
            }
            st.session_state.daten["module"].append(neues)
            try:
                speichere_daten(st.session_state.daten)
                st.success(
                    f"Modul „{f_name}“ gespeichert und dauerhaft in "
                    "data/modules.json abgelegt."
                )
            except OSError as e:
                st.warning(
                    "Modul wurde der aktuellen Sitzung hinzugefügt, konnte aber "
                    f"nicht in die Datei geschrieben werden ({e}). Auf Streamlit "
                    "Cloud ist der Schreibzugriff nicht dauerhaft — nutze dort den "
                    "Export im Tab „Verwalten & Export“."
                )

# =============================================================================
# TAB 3 – VERWALTEN & EXPORT
# =============================================================================
with tab_verwalten:
    st.subheader("Manuell erfasste Module verwalten")
    manuelle = [(i, m) for i, m in enumerate(module) if m.get("manuell")]

    if not manuelle:
        st.info("Noch keine manuell erfassten Module. Lege im Tab "
                "„Modul erfassen“ eines an.")
    else:
        for i, m in manuelle:
            c1, c2 = st.columns([5, 1])
            c1.markdown(f"**{m.get('name')}**  ·  {m.get('modul_id','')}")
            if c2.button("Löschen", key=f"del_{i}"):
                st.session_state.daten["module"].pop(i)
                try:
                    speichere_daten(st.session_state.daten)
                    st.success(f"Modul „{m.get('name')}“ gelöscht.")
                except OSError as e:
                    st.warning(f"Aus der Sitzung entfernt, Datei nicht "
                               f"geschrieben ({e}).")
                st.rerun()

    st.divider()
    st.subheader("Export")
    st.caption(
        "Lädt die komplette aktuelle Datenbasis (Excel-Module + manuell "
        "erfasste) als modules.json herunter — z. B. um sie ins GitHub-Repo "
        "zu legen, damit auch die Cloud-Version sie enthält."
    )
    st.download_button(
        "Komplette modules.json herunterladen",
        json.dumps(st.session_state.daten, ensure_ascii=False, indent=2)
            .encode("utf-8"),
        file_name="modules.json", mime="application/json",
    )
