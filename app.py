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
      /* Nur eigene Elemente einfärben – Streamlit-Standardtext NICHT anfassen,
         damit Labels/Eingaben in jedem Theme lesbar bleiben. */
      .modul-karte { background: var(--secondary-background-color, #f0f2f6);
          border: 1px solid rgba(128,128,128,0.25);
          border-radius:10px; padding:1.2rem 1.4rem; margin-bottom:1rem; }
      .badge { display:inline-block; background:#1c4966; color:#ffffff;
          border-radius:999px; padding:2px 10px; margin:2px 4px 2px 0;
          font-size:0.78rem; }
      .badge-manuell { background:#0d6b5e; color:#ffffff; }
      .sektion-titel { color:#0d6b5e; font-weight:700; margin-top:0.9rem;
          border-bottom:1px solid rgba(128,128,128,0.3); padding-bottom:2px; }
      .feld-label { opacity:0.7; font-size:0.82rem; margin-top:0.4rem; }
      .feld-wert  { font-weight:500; }
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

KEINE_ANGABE = "keine Angabe"

# Zuordnung der LISTEN-Dimensionen zu den 8 Steckbrief-Abschnitten.
# Steuert die gruppierte Filter-Startseite.
ABSCHNITT_DIMENSIONEN = {
    "1  Identifikation": [],
    "2  Zielgruppe & Adressierung": [
        "Zielgruppe", "Teilnehmerkonstellation", "Hauptzweck"],
    "3  Lernziele & Kompetenzen": [
        "Kompetenzklasse", "Fachbezogene Lerninhalte"],
    "4  Inhalt & Prozessbezug (Remanufacturing)": [
        "Abgebildete Produktlebenszyklusphase", "Prozesstyp",
        "Prozessautomatisierung", "Integrierte digitale Technologien",
        "Materialität", "Variantenanzahl", "Weitere Produktverwendung",
        "Weichensteller der Wandlungsfähigkeit"],
    "5  Didaktik & Methodik": [
        "Lernszenario", "Autonomiegrad", "Trainerrolle", "Lernaktivitätsart",
        "Standardisierung_Trainer", "Personalisierungsgrad", "Lernumgebung"],
    "6  Organisation & Einordnung": [
        "Teilnehmende pro Lernmodul", "Durchschnittsdauer Lernmodul",
        "Anzahl standardisierter Lernmodule", "Anzahl integrierter Lernmodule"],
    "7  Evaluation & Erfolgskriterien": [
        "Evaluationsebenen", "Evaluationsmethoden"],
    "8  Wirtschaftlichkeit & Trägerschaft": [
        "Betreiber", "Anschubfinanzierung", "Laufende Finanzierung",
        "Trainingsmodelle", "Schlüsselpartnerschaften", "Einrichtungskosten",
        "Betriebskosten"],
}

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
    """Modul passt, wenn einer der gewählten Werte in der normierten
    Klassifikation des Moduls steht (exakter Vergleich)."""
    if not gewaehlt:
        return True
    # Bestandsmodule: 'listen_norm'; manuell erfasste: 'listen_auswahl'
    werte = set(modul.get("listen_norm", {}).get(dim, []))
    werte |= set(modul.get("listen_auswahl", {}).get(dim, []))
    return bool(werte & set(gewaehlt))


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

    # Normierte LISTEN-Klassifikation anzeigen (Bestand: listen_norm,
    # manuell erfasst: listen_auswahl)
    norm_quelle = modul.get("listen_norm") or modul.get("listen_auswahl")
    if norm_quelle:
        st.markdown("<div class='sektion-titel'>Normierte Klassifikation "
                    "(LISTEN)</div>", unsafe_allow_html=True)
        for dim, werte in norm_quelle.items():
            werte_anzeige = [w for w in werte if w and w != KEINE_ANGABE]
            if werte_anzeige:
                st.markdown(
                    f"<div class='feld-label'>{dim}</div>"
                    f"<div class='feld-wert'>{', '.join(werte_anzeige)}</div>",
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
    # Auswahl-Zustand initialisieren
    if "filter_auswahl" not in st.session_state:
        st.session_state.filter_auswahl = {}      # {dim: [werte]}

    with st.sidebar:
        st.header("Filter")
        suchbegriff = st.text_input(
            "Freitextsuche", placeholder="z. B. YOLO, Demontage, Bloom …"
        )
        level_opt = sorted({m.get("level", "") for m in module if m.get("level")})
        gewaehlt_level = st.multiselect("Level (Schwarzer)", level_opt)
        verknuepfung = st.radio(
            "Verknüpfung der Dimensionen",
            ["UND (alle Filter erfüllt)", "ODER (mind. ein Filter)"],
        )
        st.divider()
        if st.button("Alle Filter zurücksetzen", use_container_width=True):
            st.session_state.filter_auswahl = {}
            st.rerun()

    aktive_filter = {d: w for d, w in st.session_state.filter_auswahl.items() if w}

    # --- Gruppierte Filter-Tabs (8 Abschnitte, immer sichtbar) --------------
    st.markdown("#### Filter nach Themengruppen")
    st.caption("Gruppe oben wählen; Filter darunter. Die Ergebnisse "
               "aktualisieren sich sofort.")

    # Nur Abschnitte mit Filterdimensionen erhalten einen Tab.
    filter_abschnitte = [a for a in abschnitte_reihenfolge
                         if ABSCHNITT_DIMENSIONEN.get(a)]
    # Tab-Beschriftung kompakt: Nummer + Kurzname, plus Marker bei aktiven Filtern
    def tab_label(abschnitt: str) -> str:
        dims = ABSCHNITT_DIMENSIONEN.get(abschnitt, [])
        n_aktiv = sum(1 for d in dims if st.session_state.filter_auswahl.get(d))
        kurz = abschnitt.split("  ", 1)[-1] if "  " in abschnitt else abschnitt
        nr = abschnitt.split("  ", 1)[0]
        return f"{nr}. {kurz}" + (f" ●{n_aktiv}" if n_aktiv else "")

    gruppe_tabs = st.tabs([tab_label(a) for a in filter_abschnitte])
    for tab, abschnitt in zip(gruppe_tabs, filter_abschnitte):
        with tab:
            for dim in ABSCHNITT_DIMENSIONEN[abschnitt]:
                optionen = listen.get(dim, [])
                if not optionen:
                    continue
                wahl = st.multiselect(
                    dim, optionen,
                    default=st.session_state.filter_auswahl.get(dim, []),
                    key=f"ms_{dim}",
                )
                st.session_state.filter_auswahl[dim] = wahl

    aktive_filter = {d: w for d, w in st.session_state.filter_auswahl.items()
                     if w}

    st.divider()

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
        "Die 8 Abschnitte als Tabs – du kannst frei zwischen ihnen wechseln, "
        "ohne dass Eingaben verloren gehen. Unten „Modul speichern“ schreibt "
        "dauerhaft in data/modules.json (lokal)."
    )

    e1, e2, e3, e4, e5, e6, e7, e8 = st.tabs([
        "1. Identifikation", "2. Zielgruppe", "3. Lernziele",
        "4. Inhalt & Prozess", "5. Didaktik", "6. Organisation",
        "7. Evaluation", "8. Wirtschaftlichkeit",
    ])

    with e1:
        c1, c2 = st.columns(2)
        f_id = c1.text_input("Modul-ID / Kürzel", key="e_id",
                             placeholder="z. B. M_SCHWARZER_DT_IOT")
        f_version = c2.text_input(
            "Version / Stand", key="e_version",
            value=f"V_01 / {date.today().strftime('%d.%m.%Y')}")
        f_name = st.text_input("Modulname *", key="e_name",
                               placeholder="Vollständiger Modultitel (Pflichtfeld)")
        f_autor = st.text_input("Autor:in / verantwortlich", key="e_autor",
                                placeholder="z. B. Prof. Schwarzer")
        f_ansprech = st.text_input("Ansprechpartner", key="e_ansprech",
                                   placeholder="Name der Kontaktperson")

    with e2:
        f_zielgruppe = st.multiselect("Zielgruppe(n)",
                                      listen.get("Zielgruppe", []), key="e_zg")
        f_konstellation = st.selectbox(
            "Teilnehmerkonstellation",
            [""] + listen.get("Teilnehmerkonstellation", []), key="e_konst")
        f_hauptzweck = st.multiselect("Hauptzweck",
                                      listen.get("Hauptzweck", []), key="e_zweck")
        f_vorwissen = st.text_input("Vorwissen / Voraussetzung", key="e_vor")

    with e3:
        f_lernziel = st.text_area(
            "Übergeordnetes Lernziel", key="e_lz",
            placeholder="Was können die Lernenden nach dem Modul? "
            "(„Lernende können …“)")
        f_kompetenz = st.multiselect("Kompetenzklassen (Erpenbeck)",
                                     listen.get("Kompetenzklasse", []), key="e_komp")
        f_bloom = st.text_input(
            "Kognitive Stufen (Bloom)", key="e_bloom",
            placeholder="z. B. überwiegend 3. Anwenden, 5. Bewerten")
        f_feinlernziele = st.text_area(
            "Feinlernziele", key="e_fein",
            placeholder="Einzelne, prüfbare Lernziele – je mit Bloom-Stufe")

    with e4:
        f_lebenszyklus = st.multiselect(
            "Abgebildete Produktlebenszyklusphase",
            listen.get("Abgebildete Produktlebenszyklusphase", []), key="e_lz4")
        f_lerninhalte = st.multiselect("Fachbezogene Lerninhalte",
                                       listen.get("Fachbezogene Lerninhalte", []),
                                       key="e_inh")
        f_technologien = st.multiselect(
            "Integrierte digitale Technologien",
            listen.get("Integrierte digitale Technologien", []), key="e_tech")
        f_materialitaet = st.selectbox("Materialität",
                                       [""] + listen.get("Materialität", []),
                                       key="e_mat")

    with e5:
        f_szenario = st.selectbox("Lernszenario-Strategie",
                                  [""] + listen.get("Lernszenario", []), key="e_szen")
        f_autonomie = st.selectbox("Autonomiegrad",
                                   [""] + listen.get("Autonomiegrad", []),
                                   key="e_auto")
        f_trainer = st.selectbox("Trainerrolle",
                                 [""] + listen.get("Trainerrolle", []), key="e_train")
        f_aktivitaet = st.multiselect("Lernaktivität",
                                      listen.get("Lernaktivitätsart", []), key="e_akt")

    with e6:
        c3, c4 = st.columns(2)
        f_dauer = c3.selectbox("Dauer",
                               [""] + listen.get("Durchschnittsdauer Lernmodul", []),
                               key="e_dauer")
        f_teilnehmer = c4.selectbox(
            "Teilnehmende pro Lernmodul",
            [""] + listen.get("Teilnehmende pro Lernmodul", []), key="e_tn")
        f_setting = st.selectbox("Lernumgebung",
                                 [""] + listen.get("Lernumgebung", []), key="e_set")
        f_ausstattung = st.text_area("Benötigte Ausstattung", key="e_aus")

    with e7:
        f_evalebene = st.multiselect("Evaluationsebene",
                                     listen.get("Evaluationsebenen", []), key="e_eve")
        f_evalmethode = st.multiselect("Evaluationsmethoden",
                                       listen.get("Evaluationsmethoden", []),
                                       key="e_evm")
        f_erfolg = st.text_input("Erfolgskriterium", key="e_erf")

    with e8:
        f_betreiber = st.selectbox("Betreiber",
                                   [""] + listen.get("Betreiber", []), key="e_betr")
        f_traegerschaft = st.text_input("Trägerschaft / Partner", key="e_trg")
        f_geschaeftsmodell = st.selectbox("Geschäftsmodell (Training)",
                                          [""] + listen.get("Trainingsmodelle", []),
                                          key="e_gm")
        f_partnerschaften = st.multiselect(
            "Schlüsselpartnerschaften",
            listen.get("Schlüsselpartnerschaften", []), key="e_part")
        f_einrichtungskosten = st.selectbox(
            "Einrichtungskosten",
            [""] + listen.get("Einrichtungskosten", []), key="e_kost")
        f_zustand = st.text_input("Zustand / Reifegrad", key="e_zust")

    st.divider()
    absenden = st.button("Modul speichern", type="primary", key="e_save")

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
