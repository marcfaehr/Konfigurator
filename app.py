"""
Baustein 4: Konfigurationsphase abschliessen.

Neu gegenueber Baustein 3:
  - Seitenleiste mit Fortschritt je Einheit (offene Merkmale via core.open_count)
  - Knopf "Auswertung starten", nur aktiv wenn ALLE Einheiten vollstaendig sind
    (core.is_complete). Setzt dann state["phase"] = "ergebnis".
  - Platzhalter-Ergebnisseite (die echte Berechnung kommt im naechsten Schritt).

Start:  streamlit run app.py
core.py und 4_typen_morphologie.xlsx muessen im selben Ordner liegen.
"""

import streamlit as st
import core


@st.cache_data
def lade_aus_bytes(daten_bytes):
    """Liest die Morphologie aus hochgeladenen Bytes.
    Gibt (features, types, weights) zurueck ODER loest core.GewichtFehler aus.
    Wird ueber den Inhalt zwischengespeichert (gleiche Datei -> kein Neulesen)."""
    import io
    return core.load_morphology(io.BytesIO(daten_bytes))


def init_state():
    if "state" not in st.session_state:
        st.session_state.state = core.empty_state()


def merkmal_tabelle(state, features, merkmal, weights):
    """Zeichnet EINE Merkmals-Tabelle (fuer jedes Merkmal aufgerufen).
    weights ist das Gewichts-Dict (merkmal -> Zahl) oder None (ungewichtet).
    Wird innerhalb eines Aufklappbereichs gezeichnet; der Merkmalsname steht
    bereits im Titel des Bereichs."""
    optionen = features[merkmal]

    # Rohes Gewicht klein anzeigen.
    if weights is not None:
        g = weights.get(merkmal)
        if g is not None:
            g_text = str(int(g)) if float(g).is_integer() else str(g)
            st.caption(f"Gewicht: {g_text}")
    else:
        st.caption("ungewichtet (alle Merkmale gleich)")

    breiten = [4] + [1] * len(state["units"]) + [1]

    # --- Kopfzeile ---
    kopf = st.columns(breiten)
    kopf[0].markdown("")
    for i, uid in enumerate(state["units"]):
        spalte = kopf[i + 1]
        spalte.markdown(f"**{uid}**")

        if st.session_state.get("warte_bestaetigung") == uid:
            spalte.caption("⚠ loeschen?")
            if spalte.button("Loeschen", key=f"ok_{uid}_{merkmal}", type="tertiary",
                             help=f"Einheit {uid} endgueltig loeschen", width="stretch"):
                st.session_state.zu_loeschen = uid
                st.rerun()
            if spalte.button("Abbrechen", key=f"no_{uid}_{merkmal}", type="tertiary",
                             help="Loeschen abbrechen", width="stretch"):
                st.session_state.pop("warte_bestaetigung", None)
                st.rerun()
        else:
            if spalte.button("🗑", key=f"del_{uid}_{merkmal}", help=f"Einheit {uid} loeschen"):
                st.session_state.warte_bestaetigung = uid
                st.rerun()

    if kopf[-1].button("➕", key=f"add_{merkmal}", help="Neue Betrachtungseinheit",
                       width="stretch"):
        core.add_unit(state, features)
        st.rerun()

    # --- Datenzeilen ---
    for opt in optionen:
        zeile = st.columns(breiten)
        zeile[0].write(opt)
        for i, uid in enumerate(state["units"]):
            gewaehlt = core.get_choice(state, uid, merkmal)
            ist_aktiv = (gewaehlt == opt)
            key = f"btn_{uid}_{merkmal}_{opt}"
            if zeile[i + 1].button(
                "●" if ist_aktiv else "○",
                key=key,
                type="primary" if ist_aktiv else "secondary",
                width="stretch",
            ):
                core.set_choice(state, uid, merkmal, opt)
                st.rerun()
        zeile[-1].write("")


def erfassung_tabelle(state, features, merkmal, weights):
    """Gemeinsame Erfassungstabelle fuer EIN Merkmal unter Tims Regel.
    Pro Einheit ein Welt-Dropdown (Ist-bekannt / Ist-unbekannt) und darunter
    je Auspraegung EIN Knopf, dessen Wirkung von der Welt abhaengt (core.klick_erfassung):
      - Ist-bekannt: erster Klick setzt Ist, weitere Klicks auf andere Auspraegungen
        setzen Potenzial; Klick auf das Ist setzt das Merkmal zurueck.
      - Ist-unbekannt: Klick schaltet Ausschluss an/aus (beliebig viele)."""
    optionen = features[merkmal]

    # Gewicht klein anzeigen.
    if weights is not None:
        g = weights.get(merkmal)
        if g is not None:
            g_text = str(int(g)) if float(g).is_integer() else str(g)
            st.caption(f"Gewicht: {g_text}")
    else:
        st.caption("ungewichtet (alle Merkmale gleich)")

    # Textspalte + je Einheit eine Spalte + Anlege-Spalte.
    breiten = [4] + [2] * len(state["units"]) + [1]

    # Kopfzeile: Einheitenname + Loeschen; Anlegen rechts.
    kopf = st.columns(breiten)
    kopf[0].markdown("")
    for i, uid in enumerate(state["units"]):
        spalte = kopf[1 + i]
        if st.session_state.get("warte_bestaetigung") == uid:
            spalte.markdown(f"**{uid}**  ⚠")
            c1, c2 = spalte.columns(2)
            if c1.button("Ja", key=f"ok_{uid}_{merkmal}", type="tertiary",
                         help=f"Einheit {uid} endgueltig loeschen", width="stretch"):
                st.session_state.zu_loeschen = uid
                st.rerun()
            if c2.button("Nein", key=f"no_{uid}_{merkmal}", type="tertiary",
                         help="Loeschen abbrechen", width="stretch"):
                st.session_state.pop("warte_bestaetigung", None)
                st.rerun()
        else:
            titel, knopf = spalte.columns([3, 1])
            titel.markdown(f"**{uid}**")
            if knopf.button("🗑", key=f"del_{uid}_{merkmal}", help=f"Einheit {uid} loeschen"):
                st.session_state.warte_bestaetigung = uid
                st.rerun()

    if kopf[-1].button("➕", key=f"add_{merkmal}", help="Neue Betrachtungseinheit",
                       width="stretch"):
        core.add_unit(state, features)
        st.rerun()

    # Welt-Dropdown je Einheit.
    weltzeile = st.columns(breiten)
    weltzeile[0].caption("Weiss ich das Ist?")
    welt_labels = {core.WELT_IST_BEKANNT: "Ist bekannt",
                   core.WELT_IST_UNBEKANNT: "Ist unbekannt"}
    for i, uid in enumerate(state["units"]):
        akt = core.get_welt(state, uid, merkmal)
        wahl = weltzeile[1 + i].selectbox(
            f"Welt {uid} {merkmal}",
            options=[core.WELT_IST_BEKANNT, core.WELT_IST_UNBEKANNT],
            index=0 if akt == core.WELT_IST_BEKANNT else 1,
            format_func=lambda w: welt_labels[w],
            key=f"welt_{uid}_{merkmal}",
            label_visibility="collapsed",
        )
        if wahl != akt:
            core.set_welt(state, uid, merkmal, wahl)
            st.rerun()
    weltzeile[-1].caption("")

    # Datenzeilen: je Auspraegung ein Knopf pro Einheit (Wirkung je Welt).
    for opt in optionen:
        zeile = st.columns(breiten)
        zeile[0].write(opt)
        for i, uid in enumerate(state["units"]):
            zustand = core.get_zellzustand(state, uid, merkmal, opt)
            welt = core.get_welt(state, uid, merkmal)
            label, typ = _erfassung_knopf(zustand)
            if welt == core.WELT_IST_UNBEKANNT:
                hilfe = "Ausschluss an/aus"
            elif core.get_choice(state, uid, merkmal) is core.OFFEN:
                hilfe = "Als Ist markieren"
            elif zustand == core.ZELLE_IST:
                hilfe = "Ist (Klick setzt Merkmal zurueck)"
            else:
                hilfe = "Potenzial an/aus"
            if zeile[1 + i].button(label, key=f"z_{uid}_{merkmal}_{opt}",
                                   type=typ, width="stretch", help=hilfe):
                core.klick_erfassung(state, uid, merkmal, opt)
                st.rerun()
        zeile[-1].write("")


def _erfassung_knopf(zustand):
    """Label und Knopf-Typ fuer eine Auspraegung nach ihrem Zustand."""
    if zustand == core.ZELLE_IST:
        return "✗ Ist", "primary"
    if zustand == core.ZELLE_POTENTIAL:
        return "🟠 Pot.", "secondary"
    if zustand == core.ZELLE_AUSSCHLUSS:
        return "⛔ Aus.", "secondary"
    return "○", "secondary"


def seite_erfassung(state, features, weights):
    """Schritt 1+3 vereint: gemeinsame Erfassung von Ist, Potenzial, Ausschluss."""
    st.title("Konfigurator — Zustandserfassung")

    if not state["units"]:
        st.info("Noch keine Betrachtungseinheit vorhanden. "
                "Nutze das ➕ in einer der Tabellen unten, um die erste anzulegen.")

    # Kurze Legende, damit Dropdown und Knopf-Zustaende klar sind.
    st.caption("Waehle je Einheit und Merkmal im Dropdown, ob das **Ist bekannt** "
               "ist. Bei *Ist bekannt* macht der erste Klick eine Auspraegung zum "
               "✗ Ist, weitere Klicks auf andere Auspraegungen zu 🟠 Potenzial; ein "
               "Klick auf das Ist setzt das Merkmal zurueck. Bei *Ist unbekannt* "
               "schaltet der Klick ⛔ Ausschluss an oder aus (mehrere moeglich).")

    for idx, merkmal in enumerate(features):
        sp_titel, sp_gewicht = st.columns([4, 1])
        sp_titel.subheader(merkmal)
        with sp_gewicht:
            g = st.number_input(
                "Gewicht", min_value=1.0,
                value=float(core.get_gewicht(state, merkmal)),
                step=0.5, key=f"gw_{merkmal}",
                help="Relative Wichtigkeit fuer die Aehnlichkeit: 1 = normal, "
                     "2 = doppelt so wichtig, 1,5 = anderthalbfach usw. "
                     "Gilt fuer das Merkmal ueber alle Einheiten.")
            core.set_gewicht(state, merkmal, g)
        erfassung_tabelle(state, features, merkmal, weights)
        if idx < len(features) - 1:
            st.divider()


def sidebar_navigation(state, features, types):
    """Phasenabhaengige Seitenleiste: Navigation (Weiter/Zurueck) und Hinweis,
    was zum Weiterkommen noch offen ist. Das Hochladefeld und der volle
    Erfassungsstatus erscheinen nur in Schritt 1 (Zustandserfassung)."""
    phase = state["phase"]
    if phase == "konfiguration":
        _sidebar_erfassung(state, features)
    elif phase == "ergebnis_soll":
        _sidebar_bewertung(state)
    elif phase == "detaillierung":
        _sidebar_sollfestlegung(state, features)
    elif phase == "zusammenfassung":
        _sidebar_typvergleich(state)
    elif phase == "massnahmen":
        _sidebar_massnahmen(state)


def _sidebar_bewertung(state):
    """Schritt 2 (Aehnlichkeitsbewertung): Navigation + Hinweis zur Typwahl."""
    with st.sidebar:
        st.header("Naechster Schritt")
        if st.button("← Zurueck zur Zustandserfassung", width="stretch"):
            state["phase"] = "konfiguration"
            st.rerun()
        hat_auswahl = any(core.get_engere_auswahl(state, uid)
                          for uid in state["units"])
        if st.button("Weiter zur Soll-Festlegung →", type="primary",
                     width="stretch", disabled=not hat_auswahl):
            state["phase"] = "detaillierung"
            st.rerun()
        st.divider()
        if hat_auswahl:
            st.caption("Mindestens ein Typ ist in der engeren Auswahl. Du kannst "
                       "zur Soll-Festlegung weitergehen.")
        else:
            st.caption("Waehle im Hauptbereich fuer jede Einheit bis zu drei Typen "
                       "in die engere Auswahl. Sobald mindestens ein Typ gewaehlt "
                       "ist, kannst du weiter.")


def _sidebar_sollfestlegung(state, features):
    """Schritt 3 (Soll-Festlegung): Navigation + Hinweis, was je Fall zu tun ist."""
    with st.sidebar:
        st.header("Naechster Schritt")
        if st.button("← Zurueck zur Ähnlichkeitsbewertung", width="stretch"):
            state["phase"] = "ergebnis_soll"
            st.rerun()
        # Belegt zugleich das Soll vor (Faelle 1/2/3) und prueft Vollstaendigkeit.
        vollstaendig = core.alle_detaillierungen_vollstaendig(state, features)
        if st.button("Weiter zum Typvergleich →", type="primary",
                     width="stretch", disabled=not vollstaendig):
            state["phase"] = "zusammenfassung"
            st.rerun()
        st.divider()
        st.caption("Lege je Merkmal und Typ eine Soll-Angabe fest:")
        st.caption("• Faelle 2 & 3: pruefe, welche Profilauspraegung mit welchem "
                   "Aufwand erreichbar waere, oder behalte das Ist.")
        st.caption("• Faelle 4 & 5: lege fest, welche Auspraegung mit welchem "
                   "Aufwand angestrebt wird, oder waehle „nichts anstreben“.")
        if not vollstaendig:
            st.warning("Es fehlen noch Angaben. Erst wenn ueberall eine Angabe "
                       "vorliegt, kannst du weiter.")


def _sidebar_typvergleich(state):
    """Schritt 4 (Typvergleich): Navigation + Hinweis zur Zieltyp-Festlegung."""
    with st.sidebar:
        st.header("Naechster Schritt")
        if st.button("← Zurueck zur Soll-Festlegung", width="stretch"):
            state["phase"] = "detaillierung"
            st.rerun()
        # Weiter erst, wenn jede Einheit mit engerer Auswahl einen Zieltyp hat.
        offen = [uid for uid in state["units"]
                 if core.get_engere_auswahl(state, uid)
                 and core.get_zieltyp(state, uid) is None]
        if st.button("Weiter zur Massnahmenplanung →", type="primary",
                     width="stretch", disabled=bool(offen)):
            state["phase"] = "massnahmen"
            st.rerun()
        st.divider()
        st.caption("Vergleiche die Typen der engeren Auswahl anhand des "
                   "gestaffelten Soll-Scores und lege je Einheit einen finalen "
                   "Zieltyp fest.")
        if offen:
            st.warning("Noch kein Zieltyp festgelegt fuer: " + ", ".join(offen))


def _sidebar_massnahmen(state):
    """Schritt 5 (Massnahmenplanung): Zurueck-Navigation + Hinweis."""
    with st.sidebar:
        st.header("Massnahmenplanung")
        if st.button("← Zurueck zum Typvergleich", width="stretch"):
            state["phase"] = "zusammenfassung"
            st.rerun()
        st.divider()
        st.caption("Die Handlungsfelder ergeben sich aus dem Vergleich von Ist "
                   "und Soll des Zieltyps. Ergaenze je Handlungsfeld die konkrete "
                   "Massnahme, ordne sie einer Etappe zu und benenne eine "
                   "Verantwortlichkeit.")


def _sidebar_erfassung(state, features):
    """Seitenleiste: Stand je Einheit + Auswertungsknopf."""
    with st.sidebar:
        st.header("Status")

        if not state["units"]:
            st.caption("Noch keine Betrachtungseinheit angelegt.")
            return

        gesamt = len(features)
        # Ein Merkmal ist "versorgt", wenn es entweder ein Ist hat ODER auf
        # 'Ist unbekannt' steht. Genau diese beiden Faelle erlauben das Weitergehen.
        for uid in state["units"]:
            versorgt = sum(
                1 for m in features
                if core.get_choice(state, uid, m) is not core.OFFEN
                or core.get_welt(state, uid, m) == core.WELT_IST_UNBEKANNT
            )
            offen = gesamt - versorgt
            st.progress(versorgt / gesamt,
                        text=f"Einheit {uid}: {versorgt}/{gesamt} versorgt"
                             + ("  ✓" if offen == 0 else f"  ({offen} offen)"))

        st.caption("Bei *Ist unbekannt* ist keine Ist-Angabe noetig. Bei "
                   "*Ist bekannt* muss ein Ist gesetzt sein, bevor es weitergeht. "
                   "Potenziale und Ausschluesse sind immer optional.")

        st.divider()

        # Harte Sperre: fuer jedes Merkmal mit Welt 'Ist bekannt' muss ein Ist
        # gesetzt sein. Potenziale/Ausschluesse sind optional.
        probleme = core.erfassung_unvollstaendig(state, features)

        if st.button("Auswertung starten", type="primary", width="stretch",
                     disabled=bool(probleme)):
            state["phase"] = "ergebnis_soll"
            st.rerun()

        if probleme:
            # Nach Einheit gruppieren fuer eine kompakte Meldung.
            pro_einheit = {}
            for uid, merkmal, _ in probleme:
                pro_einheit.setdefault(uid, []).append(merkmal)
            zeilen = "\n".join(
                f"- Betrachtungseinheit {uid}: {len(ms)} Merkmal(e) mit "
                f"'Ist bekannt' aber ohne Ist"
                for uid, ms in pro_einheit.items()
            )
            st.warning(
                "Bevor es weitergeht, fehlt noch etwas:\n\n" + zeilen
                + "\n\nSetze dort entweder ein Ist, oder stelle das Merkmal auf "
                "'Ist unbekannt', wenn du den heutigen Zustand nicht kennst."
            )


def seite_konfiguration(state, features, weights):
    st.title("Konfigurator — Ist-Aufnahme")

    if not state["units"]:
        st.info("Noch keine Betrachtungseinheit vorhanden. "
                "Nutze das ➕ in einer der Tabellen unten, um die erste anzulegen.")

    # Alle Merkmalstabellen direkt untereinander (ohne Aufklappen).
    for idx, merkmal in enumerate(features):
        st.subheader(merkmal)
        merkmal_tabelle(state, features, merkmal, weights)
        if idx < len(features) - 1:
            st.divider()


def _fmt(x):
    """Zahl als Prozent formatieren, None als '—'."""
    if x is None:
        return "—"
    return f"{x*100:.0f}%"


def _status_anzeige(status):
    """Wandelt einen core.STATUS_*-Wert in (Symbol, Klartext) fuer die Anzeige."""
    return {
        core.STATUS_IST:           ("✅", "Passt im Ist"),
        core.STATUS_POTENTIAL:     ("🟠", "Passt im Potenzial"),
        core.STATUS_IST_UNPASSEND: ("🔶", "Passt weder im Ist noch im Potenzial"),
        core.STATUS_OFFEN:         ("⬜", "Offen / Nicht blockiert"),
        core.STATUS_BLOCKIERT:     ("⛔", "Ausgeschlossen / Blockiert"),
    }.get(status, ("", status))


def potential_tabelle(state, features, merkmal):
    """Zeichnet EINE Potential/Ausschluss-Tabelle (Phase 3) fuer ein Merkmal.
    Gleiche Struktur wie die Ist-Tabelle, aber:
      - keine neuen Einheiten, kein Loeschen
      - Ist-Auspraegung wird als festes Kreuz angezeigt (kein Bedienelement)
      - fuer jede andere Auspraegung ein Dropdown: '', 'Potential', 'Ausschluss'
      - 'keine Angabe' bekommt kein Dropdown (sinnlos als Potential/Ausschluss)"""
    optionen = features[merkmal]

    # Spalten: breite Textspalte + je Einheit eine Spalte (kein "+"-Knopf hier).
    breiten = [4] + [2] * len(state["units"])

    # Kopfzeile mit Einheiten-Buchstaben.
    kopf = st.columns(breiten)
    kopf[0].markdown("")
    for i, uid in enumerate(state["units"]):
        kopf[i + 1].markdown(f"**{uid}**")

    # Dropdown-Optionen: intern "", "potential", "ausschluss";
    # angezeigt als leeres Feld, "Potential", "Ausschluss".
    status_optionen = ["", "potential", "ausschluss"]
    def status_label(s):
        return {"": "—", "potential": "Potenzial", "ausschluss": "Ausschluss"}[s]

    for opt in optionen:
        zeile = st.columns(breiten)
        zeile[0].write(opt)
        for i, uid in enumerate(state["units"]):
            zelle = zeile[i + 1]
            ist_wert = core.get_choice(state, uid, merkmal)

            # 1) Ist-Auspraegung: festes Kreuz, kein Bedienelement.
            if opt == ist_wert:
                zelle.markdown("✗ &nbsp;*(Ist)*", unsafe_allow_html=True)
                continue
            # 2) Alle anderen: Dropdown mit drei Zustaenden.
            aktueller = core.get_pa_status(state, uid, merkmal, opt)
            idx = status_optionen.index(aktueller)
            key = f"pa_{uid}_{merkmal}_{opt}"
            wahl = zelle.selectbox(
                label=f"{uid} – {opt}",
                options=status_optionen,
                index=idx,
                format_func=status_label,
                key=key,
                label_visibility="collapsed",
            )
            if wahl != aktueller:
                core.set_pa_status(state, uid, merkmal, opt, wahl)
                st.rerun()


def seite_potential(state, features):
    st.title("Konfigurator — Potenzial-Aufnahme")

    col_zur, col_weiter, col_info = st.columns([1, 1, 2])
    with col_zur:
        if st.button("← Zurueck zum Ist-Ergebnis"):
            state["phase"] = "ergebnis"
            st.rerun()
    with col_weiter:
        if st.button("Weiter zum Potenzial-Ergebnis →", type="primary"):
            state["phase"] = "ergebnis_soll"
            st.rerun()
    with col_info:
        st.caption("Mehrere Potenzial-/Ausschluss-Auspraegungen moeglich. "
                   "Das Ist (✗) steht fest.")

    if not state["units"]:
        st.info("Keine Betrachtungseinheiten vorhanden.")
        return

    st.divider()
    for idx, merkmal in enumerate(features):
        st.subheader(merkmal)
        potential_tabelle(state, features, merkmal)
        if idx < len(features) - 1:
            st.divider()


def seite_ergebnis_ist(state, features, types, weights):
    """ERSTES Dashboard (Phase 'ergebnis'): zeigt NUR den Ist-Zustand.
    Kein Soll, keine harten Verstoesse (Potential ist hier noch nicht erfasst)."""
    import pandas as pd
    import altair as alt
    FARBE_IST = "#1f77b4"   # Blau (durchgaengig fuer Ist)

    st.title("Konfigurator — Ist-Ergebnis")

    col_zurueck, col_weiter, _ = st.columns([1, 1, 3])
    with col_zurueck:
        if st.button("← Zurueck zur Konfiguration"):
            state["phase"] = "konfiguration"
            st.rerun()
    with col_weiter:
        if st.button("Weiter zur Potenzial-Aufnahme →", type="primary"):
            state["phase"] = "potential"
            st.rerun()

    if not state["units"]:
        st.info("Keine Betrachtungseinheiten vorhanden.")
        return

    for uid in state["units"]:
        auswahl = state["matrix"][uid]
        rang_ist = core.ranking(auswahl, types, features, weights)
        best_name, best_ist = rang_ist[0]
        eind = core.eindeutigkeit(rang_ist)
        anteil = core.abdeckung_anteil(auswahl, features, weights)

        st.header(f"Betrachtungseinheit {uid}")

        # Kopf-Kennzahlen: nur Ist.
        k1, k2, k3 = st.columns(3)
        k1.metric("Bestpassender Typ", best_name)
        k2.metric("Ist-Uebereinstimmung", _fmt(best_ist))
        k3.metric("Eindeutigkeit (Ist)", _fmt(eind))

        # Balkendiagramm: nur Ist, in Blau.
        st.subheader("Rangfolge der Typen — Ist")
        df_chart = pd.DataFrame(
            [{"Typ": n, "Wert": k * 100} for n, k in rang_ist]
        )
        typ_reihenfolge = [n for n, _ in rang_ist]
        chart = (
            alt.Chart(df_chart)
            .mark_bar(color=FARBE_IST)
            .encode(
                x=alt.X("Typ:N", sort=typ_reihenfolge, title=None),
                y=alt.Y("Wert:Q", title="Übereinstimmung (%)",
                        scale=alt.Scale(domain=[0, 100])),
                tooltip=["Typ", alt.Tooltip("Wert:Q", format=".0f")],
            )
            .properties(height=260)
        )
        st.altair_chart(chart, use_container_width=True)

        # Tabelle: nur Ist + Ist-ueber-beantwortete.
        kpi_beantwortet = {
            n: core.uebereinstimmung_beantwortet(auswahl, types[n], features, weights)
            for n, _ in rang_ist
        }
        df_tab = pd.DataFrame([
            {"Typ": n, "Ist": _fmt(k), "nur beantwortete": _fmt(kpi_beantwortet[n])}
            for n, k in rang_ist
        ])
        st.dataframe(df_tab, hide_index=True, width="stretch")

        # Guete: Abdeckung + Ist-ueber-beantwortete des Bestpassers.
        st.subheader("Guete der Aussage")
        g1, g2 = st.columns(2)
        g1.metric("Abdeckung (beantwortete Merkmale)", _fmt(anteil),
                  help="Gewichteter Anteil der Merkmale, die nicht 'keine Angabe' sind.")
        g2.metric(f"Ist {best_name} (nur beantwortete)",
                  _fmt(kpi_beantwortet[best_name]),
                  help="Ist-KPI, aber nur ueber beantwortete Merkmale gerechnet.")
        if anteil < 1.0:
            st.caption("Diese Einheit hat 'keine Angabe'-Merkmale; der Ist-KPI ist "
                       "dadurch niedriger als der Wert ueber nur beantwortete Merkmale.")

        st.divider()


def seite_ergebnis_soll(state, features, types, weights):
    """Dashboard (Phase 'ergebnis_soll'): dreiteiliges Aehnlichkeitsmass je Typ
    als Intervall (min..max) mit bereinigtem Wert, fuer Ist (blau) und
    Potenzial (orange) uebereinander. Zahlen einklappbar, Guete-KPIs,
    Status je Typ und Typfestlegung."""
    import pandas as pd
    import altair as alt
    FARBE_IST = "#1f77b4"    # Blau
    FARBE_SOLL = "#ff7f0e"   # Orange

    st.title("Konfigurator — Ähnlichkeitsbewertung")

    if not state["units"]:
        st.info("Keine Betrachtungseinheiten vorhanden.")
        return

    for uid in state["units"]:
        auswahl = state["matrix"][uid]
        potential_unit = state["potential"][uid]
        ausschluss_unit = state["ausschluss"][uid]

        # Alle sechs Werte je Typ berechnen (Ist + Soll, je min/bereinigt/max).
        werte = {}
        for name, profil in types.items():
            ist_min = core.uebereinstimmung(auswahl, profil, features, weights, ausschluss_unit)
            ist_max = core.uebereinstimmung_max(auswahl, profil, features, weights, ausschluss_unit)
            ist_ber = core.uebereinstimmung_beantwortet(auswahl, profil, features, weights, ausschluss_unit)
            soll_min = core.uebereinstimmung_soll(auswahl, potential_unit, profil, features, weights, ausschluss_unit)
            soll_max = core.uebereinstimmung_soll_max(auswahl, potential_unit, profil, features, weights, ausschluss_unit)
            soll_ber = core.uebereinstimmung_soll_bereinigt(auswahl, potential_unit, profil, features, weights, ausschluss_unit)
            werte[name] = {
                "ist_min": ist_min, "ist_max": ist_max, "ist_ber": ist_ber,
                "soll_min": soll_min, "soll_max": soll_max, "soll_ber": soll_ber,
            }

        # Rangfolge nach dem BEREINIGTEN Ist-Wert (faire Schaetzung).
        # None (nichts beantwortet) ans Ende sortieren.
        def sortkey(n):
            b = werte[n]["ist_ber"]
            return (b is not None, b if b is not None else 0)
        typ_reihenfolge = sorted(types.keys(), key=sortkey, reverse=True)
        best_name = typ_reihenfolge[0]

        # Eindeutigkeit auf Basis des bereinigten Ist-Werts (Abstand 1./2.).
        ber_sorted = sorted(
            [werte[n]["ist_ber"] or 0 for n in typ_reihenfolge], reverse=True)
        eind_ist = (ber_sorted[0] - ber_sorted[1]) if len(ber_sorted) >= 2 else None
        # Eindeutigkeit Potenzial (bereinigt).
        ber_soll_sorted = sorted(
            [werte[n]["soll_ber"] or 0 for n in typ_reihenfolge], reverse=True)
        eind_soll = (ber_soll_sorted[0] - ber_soll_sorted[1]) if len(ber_soll_sorted) >= 2 else None

        anteil = core.abdeckung_anteil(auswahl, features, weights)
        hv_best = core.harte_verstoesse(ausschluss_unit, types[best_name], features)

        st.header(f"Betrachtungseinheit {uid}")

        # Bestpassender Typ getrennt fuer Ist und Potenzial (bereinigt).
        # Bei Gleichstand werden ALLE Typen mit dem Hoechstwert genannt.
        def _beste(schluessel):
            vorhanden = [(n, werte[n][schluessel]) for n in typ_reihenfolge
                         if werte[n][schluessel] is not None]
            if not vorhanden:
                return "—", None
            max_wert = max(w for _, w in vorhanden)
            namen = [n for n, w in vorhanden if abs(w - max_wert) < 1e-9]
            return ", ".join(namen), max_wert

        best_ist_namen, best_ist_wert = _beste("ist_ber")
        best_pot_namen, best_pot_wert = _beste("soll_ber")

        # Zeile 1: Ist
        z1a, z1b = st.columns([2, 1])
        z1a.metric("Bestpassender Typ (Ist, bereinigt)", best_ist_namen)
        z1b.metric("Ist bereinigt",
                   _fmt(best_ist_wert) if best_ist_wert is not None else "—")
        # Zeile 2: Potenzial
        z2a, z2b = st.columns([2, 1])
        z2a.metric("Bestpassender Typ (Potenzial, bereinigt)", best_pot_namen)
        z2b.metric("Potenzial bereinigt",
                   _fmt(best_pot_wert) if best_pot_wert is not None else "—")

        # --- Intervall-Diagramm: pro Typ Ist (blau) und Soll (orange) ---
        st.subheader("Aehnlichkeit je Typ — Intervall (min bis max)")
        st.caption("Balken = Spanne von Minimum bis Maximum · Strich = bereinigter "
                   "Wert · blau = Ist · orange = Potenzial. Schmaler Balken = "
                   "wenig offen, breiter Balken = viel offen.")

        balken = []   # Intervall-Balken (min..max)
        marker = []   # bereinigte Werte als Striche
        for name in typ_reihenfolge:
            w = werte[name]
            balken.append({"Typ": name, "Art": "Ist",
                           "min": w["ist_min"] * 100, "max": w["ist_max"] * 100})
            balken.append({"Typ": name, "Art": "Potenzial",
                           "min": w["soll_min"] * 100, "max": w["soll_max"] * 100})
            if w["ist_ber"] is not None:
                marker.append({"Typ": name, "Art": "Ist", "wert": w["ist_ber"] * 100})
            if w["soll_ber"] is not None:
                marker.append({"Typ": name, "Art": "Potenzial", "wert": w["soll_ber"] * 100})

        df_balken = pd.DataFrame(balken)
        df_marker = pd.DataFrame(marker)
        farb_skala = alt.Scale(domain=["Ist", "Potenzial"], range=[FARBE_IST, FARBE_SOLL])

        # Intervall-Balken als horizontale Bereiche, je Typ zwei Zeilen (Ist/Potenzial).
        bars = (
            alt.Chart(df_balken)
            .mark_bar(height=12, cornerRadius=3)
            .encode(
                y=alt.Y("Typ:N", sort=typ_reihenfolge, title=None),
                yOffset=alt.YOffset("Art:N", sort=["Ist", "Potenzial"]),
                x=alt.X("min:Q", title="Aehnlichkeit (%)", scale=alt.Scale(domain=[0, 100])),
                x2="max:Q",
                color=alt.Color("Art:N", scale=farb_skala, title=None),
                tooltip=["Typ", "Art",
                         alt.Tooltip("min:Q", title="Minimum", format=".0f"),
                         alt.Tooltip("max:Q", title="Maximum", format=".0f")],
            )
        )
        # Bereinigte Werte als kurze senkrechte Striche (Tick).
        ticks = (
            alt.Chart(df_marker)
            .mark_tick(thickness=2.5, size=16, color="#042C53")
            .encode(
                y=alt.Y("Typ:N", sort=typ_reihenfolge),
                yOffset=alt.YOffset("Art:N", sort=["Ist", "Potenzial"]),
                x=alt.X("wert:Q"),
                tooltip=[alt.Tooltip("wert:Q", title="bereinigt", format=".0f")],
            )
        )
        chart = (bars + ticks).properties(height=max(180, 60 * len(typ_reihenfolge)))
        st.altair_chart(chart, use_container_width=True)

        # --- Zahlen einklappbar darunter ---
        with st.expander("Genaue Zahlen (alle drei Maße, Ist und Potenzial)", expanded=False):
            zeilen = []
            for name in typ_reihenfolge:
                w = werte[name]
                hv = core.harte_verstoesse(ausschluss_unit, types[name], features)
                zeilen.append({
                    "Typ": name,
                    "Ist min": _fmt(w["ist_min"]),
                    "Ist berein.": _fmt(w["ist_ber"]),
                    "Ist max": _fmt(w["ist_max"]),
                    "Potenzial min": _fmt(w["soll_min"]),
                    "Potenzial berein.": _fmt(w["soll_ber"]),
                    "Potenzial max": _fmt(w["soll_max"]),
                    "harte V.": hv,
                })
            st.dataframe(pd.DataFrame(zeilen), hide_index=True, width="stretch")

        # --- Guete-Kennzahlen ---
        st.subheader("Guete der Aussage")
        g1, g2, g3 = st.columns(3)
        g1.metric("Abdeckung (Ist)", _fmt(anteil),
                  help="Gewichteter Anteil der Merkmale mit Ist-Angabe.")
        g2.metric("Eindeutigkeit (Ist)", _fmt(eind_ist),
                  help="Abstand bester zu zweitbester Typ, bereinigter Ist-Wert.")
        g3.metric("Eindeutigkeit (Potenzial)", _fmt(eind_soll),
                  help="Abstand bester zu zweitbester Typ, bereinigter Potenzial-Wert.")

        # Fuer den hinteren (unveraenderten) Teil benoetigte Groessen bereitstellen.
        rang_ist = [(n, werte[n]["ist_min"]) for n in typ_reihenfolge]
        soll_map = {n: (werte[n]["soll_ber"] if werte[n]["soll_ber"] is not None
                        else werte[n]["soll_min"]) for n in typ_reihenfolge}

        hinweise = []
        if anteil < 1.0:
            hinweise.append("Diese Einheit hat Merkmale ohne Ist-Angabe; deshalb "
                            "spannen die Ist-Balken ein Intervall auf.")
        if hv_best > 0:
            hinweise.append(f"Achtung: {hv_best} Merkmal(e) koennen den Typ {best_name} "
                            "strukturell nie erreichen (harte Verstoesse durch Ausschluss).")
        for h in hinweise:
            st.caption(h)

        # --- Merkmals-Status je Typ (einklappbar, standardmaessig zu) ---
        st.subheader("Merkmals-Status je Typ")
        st.caption("Aufklappen, um pro Typ zu sehen, welche Merkmale im Ist passen, "
                   "über Potenzial erreichbar sind, im Ist nicht passen, noch offen "
                   "oder blockiert sind.")
        for name, _ in rang_ist:
            ist_kpi = dict(rang_ist)[name]
            soll_kpi = soll_map[name]
            titel = (f"{name}   ·   Ist {_fmt(ist_kpi)} / Potenzial {_fmt(soll_kpi)}")
            with st.expander(titel, expanded=False):
                zeilen_status = []
                for m in features:
                    status = core.merkmal_status(
                        auswahl, potential_unit, ausschluss_unit, types[name], m)
                    symbol, klartext = _status_anzeige(status)
                    zeilen_status.append({
                        "Merkmal": m,
                        "Status": f"{symbol} {klartext}",
                    })
                st.dataframe(pd.DataFrame(zeilen_status), hide_index=True, width="stretch")

        # --- Engere Auswahl fuer die Detaillierung (bis zu 3 Typen) ---
        st.subheader("Engere Auswahl")
        typ_namen = [n for n, _ in rang_ist]
        aktuelle_auswahl = core.get_engere_auswahl(state, uid)
        wahl = st.multiselect(
            f"Typen für Betrachtungseinheit {uid} "
            "(werden anschliessend ausgestaltet und verglichen)",
            options=typ_namen,
            default=[t for t in aktuelle_auswahl if t in typ_namen],
            key=f"engere_auswahl_{uid}",
            help="Je mehr Typen gewaehlt sind, desto aufwaendiger wird die "
                 "Soll-Festlegung, da dort jedes Merkmal je Typ einzeln zu "
                 "entscheiden ist. Eine feste Obergrenze gibt es bewusst nicht.",
        )
        if wahl != aktuelle_auswahl:
            # Differenz anwenden: neue aufnehmen (legt Soll an), entfernte loeschen
            # (raeumt Soll und ggf. finalen Zieltyp auf).
            for t in wahl:
                if t not in aktuelle_auswahl:
                    core.add_to_engere_auswahl(state, uid, t)
            for t in list(aktuelle_auswahl):
                if t not in wahl:
                    core.remove_from_engere_auswahl(state, uid, t)
            st.rerun()

        if not wahl:
            st.caption("Noch keine Typen ausgewaehlt. Waehle die Typen, die du in der "
                       "Soll-Festlegung ausgestalten und im Typvergleich vergleichen moechtest.")
        else:
            # Kurze Rueckmeldung je gewaehltem Typ, harte Verstoesse als Warnsignal.
            teile = []
            for t in wahl:
                hv = core.harte_verstoesse(ausschluss_unit, types[t], features)
                marke = f" ⛔{hv}" if hv > 0 else ""
                teile.append(f"**{t}** (Potenzial {_fmt(soll_map[t])}{marke})")
            st.caption("In der engeren Auswahl: " + " · ".join(teile)
                       + ". Der finale Zieltyp wird im Typvergleich festgelegt.")

        st.divider()


def detail_merkmal_block(state, features, types, uid, merkmal, typ, status):
    """Merkmals-Tabelle in der Detaillierung, bezogen auf EINEN Typ (Tab).
    Knopfregel nach Wissenssituation:
      - Fall 1 (Ist passt): kein Knopf, das Soll ist automatisch das Ist.
      - Fall 2 & 3 (es gibt ein Ist): Knoepfe hinter den Kandidaten (Ist,
        Potenziale, Profil). Das Ist ist initial angestrebt; der Nutzer kann es
        belassen oder zu einem Potenzial oder einer Profilauspraegung wechseln.
      - Fall 4 & 5 (kein Ist): Knoepfe hinter ALLEN Auspraegungen, dazu der Knopf
        'nichts anstreben'. Der Nutzer muss aktiv eine Auspraegung oder 'nichts
        anstreben' waehlen. Ausgeschlossene Auspraegungen sind waehlbar, ohne den
        Ausschluss in Schritt 1 zu aendern (keine Rueckkopplung).
    Ist eine echte Auspraegung angestrebt, die vom Ist abweicht, ist der Aufwand
    Pflicht.
    """
    profil = types[typ].get(merkmal, set())
    optionen = features[merkmal]
    ist_wert = core.get_choice(state, uid, merkmal)
    soll_wert = core.get_soll(state, uid, typ, merkmal)

    ist_unbekannt_fall = status in (core.STATUS_OFFEN, core.STATUS_BLOCKIERT)

    # Kandidaten (Auspraegungen mit Knopf) fallabhaengig bestimmen (siehe
    # core.soll_kandidaten): Fall 2 Ist + passende Potenziale, Fall 3 Ist +
    # Profil, Fall 4/5 alle, Fall 1 keine.
    kandidaten = core.soll_kandidaten(state, uid, types[typ], merkmal,
                                      status, optionen)

    # (Die Soll-Vorbelegung fuer Faelle 1/2/3 erfolgt zentral in
    #  seite_detaillierung via core.soll_vorbelegen, damit auch nicht geoeffnete
    #  Tabs vollstaendig sind.)

    st.markdown(f"**{merkmal}**")

    breiten = [1, 6, 3, 3]
    kopf = st.columns(breiten)
    kopf[0].caption("Ziel")
    kopf[1].caption("Auspraegung")
    kopf[2].caption("deine Angabe")
    kopf[3].caption("Soll")

    for opt in optionen:
        z = st.columns(breiten)

        # Spalte 1: Profil-Markierung.
        z[0].markdown("✅" if opt in profil else "&nbsp;", unsafe_allow_html=True)

        # Spalte 2: Auspraegungstext, erlaubte fett.
        if opt in profil:
            z[1].markdown(f"**{opt}**")
        else:
            z[1].write(opt)

        # Spalte 3: Herkunft aus den Vorphasen.
        pa = core.get_pa_status(state, uid, merkmal, opt)
        if opt == ist_wert:
            z[2].markdown("✗ *(Ist)*", unsafe_allow_html=True)
        elif pa == "potential":
            z[2].markdown("🟠 Potenzial")
        elif pa == "ausschluss":
            z[2].markdown("⛔ Ausschluss")
        else:
            z[2].markdown("&nbsp;", unsafe_allow_html=True)

        # Spalte 4: Soll-Auspraegung.
        _detail_soll_spalte(state, uid, typ, merkmal, opt, status, soll_wert,
                            ist_wert, kandidaten, z[3])

    # In Fall 4 & 5: zusaetzlich der Knopf 'nichts anstreben'.
    soll_final = core.get_soll(state, uid, typ, merkmal)
    if ist_unbekannt_fall:
        _nichts_anstreben_knopf(state, uid, typ, merkmal, soll_final)

    # Aufwand ist Pflicht, sobald eine echte Auspraegung angestrebt ist, die vom
    # Ist abweicht. Bei belassenem Ist oder 'nichts anstreben' entfaellt er.
    if soll_final not in (core.OFFEN, core.NICHTS_ANSTREBEN) and soll_final != ist_wert:
        _aufwand_auswahl(state, uid, typ, merkmal, soll_final)


def _detail_soll_spalte(state, uid, typ, merkmal, opt, status, soll_wert,
                        ist_wert, kandidaten, spalte):
    """Rechte Spalte (angestrebt) einer Auspraegungszeile. Fall 1: kein Knopf,
    das Ist ist angestrebt. Sonst: Knopf nur fuer Kandidaten. Ausgeschlossene
    Auspraegungen koennen (in Fall 4/5) gewaehlt werden, ohne den Ausschluss in
    Schritt 1 zu aendern."""
    if status == core.STATUS_IST:
        if opt == ist_wert:
            spalte.markdown("● **Soll**")
        else:
            spalte.markdown("&nbsp;", unsafe_allow_html=True)
        return

    if opt not in kandidaten:
        spalte.markdown("&nbsp;", unsafe_allow_html=True)
        return

    _soll_knopf(state, uid, typ, merkmal, opt, soll_wert, spalte)


def _aufwand_auswahl(state, uid, typ, merkmal, soll_wert):
    """Pflicht-Aufwandsabfrage (gering/mittel/hoch) fuer ein Merkmal mit
    gewaehltem Soll, das vom Ist abweicht. Der Wert misst, wie aufwaendig es ist,
    das Merkmal vom heutigen Zustand in die Soll-Auspraegung zu bringen."""
    aktuell = core.get_aufwand(state, uid, typ, merkmal)
    sp = st.columns([3, 2, 2, 2])
    if aktuell is None:
        sp[0].markdown(f"⚠️ **Aufwand für „{soll_wert}“** (Pflicht)")
    else:
        sp[0].markdown(f"Aufwand für „{soll_wert}“")
    for i, stufe in enumerate(core.AUFWAND_STUFEN):
        aktiv = (aktuell == stufe)
        if sp[i + 1].button(core.AUFWAND_LABEL[stufe],
                            key=f"aufw_{uid}_{typ}_{merkmal}_{stufe}",
                            type="primary" if aktiv else "secondary",
                            width="stretch"):
            core.set_aufwand(state, uid, typ, merkmal, stufe)
            st.rerun()


def _soll_knopf(state, uid, typ, merkmal, opt, soll_wert, spalte):
    """Wahl-Knopf fuer die Soll-Auspraegung, fuer EINEN Typ. Kein Toggle:
    die aktive Auspraegung ist ein Marker, inaktive sind waehlbare Knoepfe. Der
    Key enthaelt den Typ, damit Knoepfe verschiedener Tabs nicht kollidieren."""
    if opt == soll_wert:
        spalte.markdown("● **Soll**")
    else:
        if spalte.button("○ waehlen", key=f"soll_{uid}_{typ}_{merkmal}_{opt}",
                         type="secondary", width="stretch"):
            core.set_soll(state, uid, typ, merkmal, opt)
            st.rerun()


def _nichts_anstreben_knopf(state, uid, typ, merkmal, soll_wert):
    """Knopf 'nichts anstreben' fuer die Faelle 4 und 5 (kein Ist). Eine bewusste
    Entscheidung, fuer dieses Merkmal zu diesem Typ nichts anzustreben; sie zaehlt
    im Soll-Score als Nichttreffer und erfuellt die Pflicht zur Angabe. Toggle:
    erneuter Klick macht sie rueckgaengig (zurueck zu offen/unentschieden)."""
    if soll_wert == core.NICHTS_ANSTREBEN:
        if st.button("● nichts anstreben (gewählt)",
                     key=f"nichts_{uid}_{typ}_{merkmal}", type="primary"):
            core.set_soll(state, uid, typ, merkmal, core.OFFEN)
            st.rerun()
    else:
        if st.button("nichts anstreben",
                     key=f"nichts_{uid}_{typ}_{merkmal}", type="secondary"):
            core.set_soll(state, uid, typ, merkmal, core.NICHTS_ANSTREBEN)
            st.rerun()


def _detail_tab_inhalt(state, features, types, uid, typ):
    """Inhalt eines Typ-Tabs in der Detaillierung: Status-Kennzahlen und die fuenf
    Abschnitte mit Merkmalstabellen, bezogen auf DIESEN Typ. Die Festlegung des
    finalen Zieltyps erfolgt spaeter in der Zusammenfassung."""
    # --- Merkmale nach Status gruppieren (fuer DIESEN Typ) ---
    gruppen = {core.STATUS_IST: [], core.STATUS_POTENTIAL: [],
               core.STATUS_IST_UNPASSEND: [], core.STATUS_OFFEN: [],
               core.STATUS_BLOCKIERT: []}
    for m in features:
        s = core.merkmal_status(state["matrix"][uid], state["potential"][uid],
                                state["ausschluss"][uid], types[typ], m)
        gruppen[s].append(m)
    anzahl = {k: len(v) for k, v in gruppen.items()}

    # --- Statuszusammenfassung als Kennzahlen (fuenf Faelle) ---
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("✅ Passt im Ist", anzahl[core.STATUS_IST])
    k2.metric("🟠 Passt im Potenzial", anzahl[core.STATUS_POTENTIAL])
    k3.metric("🔶 Weder Ist noch Potenzial", anzahl[core.STATUS_IST_UNPASSEND])
    k4.metric("⬜ Offen / Nicht blockiert", anzahl[core.STATUS_OFFEN])
    k5.metric("⛔ Ausgeschlossen / Blockiert", anzahl[core.STATUS_BLOCKIERT])

    # --- Fuenf aufklappbare Abschnitte, je Status einer ---
    abschnitte = [
        (core.STATUS_IST,       "✅ Passt im Ist",
         "Die gewaehlte Ist-Auspraegung liegt im Profil und wird angestrebt. "
         "Hier besteht kein Handlungsbedarf und kein Aufwand."),
        (core.STATUS_POTENTIAL, "🟠 Passt im Potenzial",
         "Eine als Potenzial erfasste Auspraegung liegt im Profil. Waehle die "
         "Soll-Auspraegung und gib den Aufwand an, das Potenzial auszuschoepfen."),
        (core.STATUS_IST_UNPASSEND, "🔶 Passt weder im Ist noch im Potenzial",
         "Weder die Ist-Auspraegung noch eine Potenzial-Auspraegung liegt im Profil. "
         "Entweder du behaeltst das Ist und akzeptierst, dass die Einheit hier nicht "
         "zum Typ passt, oder du strebst eine profilkonforme Auspraegung an und gibst "
         "den Aufwand an, vom heutigen Ist dorthin zu gelangen."),
        (core.STATUS_OFFEN,     "⬜ Offen / Nicht blockiert",
         "Fuer dieses Merkmal liegt keine Angabe vor. Du hast die freie Wahl: "
         "waehle eine Soll-Auspraegung (orientiere dich am ✅ Profil) und "
         "gib den Aufwand an, sie zu etablieren, oder waehle „nichts anstreben“."),
        (core.STATUS_BLOCKIERT, "⛔ Ausgeschlossen / Blockiert",
         "Alle profilkonformen Auspraegungen wurden ausgeschlossen. Du hast "
         "dennoch die freie Wahl: auch eine ausgeschlossene Auspraegung ist "
         "waehlbar, ohne die Erfassung in Schritt 1 zu aendern. Waehle eine "
         "Soll-Auspraegung und gib den Aufwand an, oder waehle „nichts "
         "anstreben“."),
    ]
    for status, titel, hinweis in abschnitte:
        merkmale = gruppen[status]
        with st.expander(f"{titel}  ({len(merkmale)})",
                         expanded=(status in (core.STATUS_POTENTIAL,
                                              core.STATUS_IST_UNPASSEND,
                                              core.STATUS_OFFEN)
                                   and len(merkmale) > 0)):
            if not merkmale:
                st.caption("— keine Merkmale in diesem Status —")
            else:
                st.caption(hinweis)
                for m in merkmale:
                    with st.container(border=True):
                        detail_merkmal_block(state, features, types, uid, m,
                                             typ, status)


def seite_detaillierung(state, features, types, weights):
    st.title("Konfigurator — Soll-Festlegung")

    # Das Soll wird in der Seitenleiste (alle_detaillierungen_vollstaendig) fuer
    # die Faelle 1/2/3 vorbelegt; die Seitenleiste laeuft vor dieser Seite.
    if not state["units"]:
        st.info("Keine Betrachtungseinheiten vorhanden.")
        return

    # Nur Einheiten mit mindestens einem Typ in der engeren Auswahl koennen
    # detailliert werden.
    ohne_auswahl = [uid for uid in state["units"]
                    if not core.get_engere_auswahl(state, uid)]
    if ohne_auswahl:
        st.warning("Fuer folgende Einheiten ist noch kein Typ in der engeren "
                   f"Auswahl: {', '.join(ohne_auswahl)}. Waehle die zu "
                   "vergleichenden Typen im Potenzial-Ergebnis aus.")

    st.divider()
    for uid in state["units"]:
        st.header(f"Betrachtungseinheit {uid}")
        auswahl = core.get_engere_auswahl(state, uid)
        if not auswahl:
            st.caption("Keine Typen in der engeren Auswahl — uebersprungen.")
            st.divider()
            continue

        final = core.get_zieltyp(state, uid)
        # Ein Tab je Typ der engeren Auswahl; der finale Zieltyp ist markiert.
        tab_titel = [(f"⭐ {t}" if t == final else t) for t in auswahl]
        for tab, typ in zip(st.tabs(tab_titel), auswahl):
            with tab:
                _detail_tab_inhalt(state, features, types, uid, typ)

        st.divider()

def _stufendiagramm(state, features, types, weights, uid, auswahl):
    """Treppendiagramm je Betrachtungseinheit: erreichbare Aehnlichkeit ueber die
    vier Aufwandsstufen k = 0..3, ein Verlauf je Typ der engeren Auswahl.

    Die Aehnlichkeit ist ausschliesslich an den vier Stufen definiert und bleibt
    dazwischen unveraendert, daher die Treppendarstellung (step-after) statt einer
    durchgezogenen Verbindung."""
    import pandas as pd

    zeilen = []
    for typ in auswahl:
        score = core.soll_score_gestaffelt(state, uid, typ, types[typ],
                                           features, weights)
        for k in core.AUFWAND_STUFEN_K:
            zeilen.append({"stufe": k, "aehnlichkeit": score[k], "typ": typ})
    df = pd.DataFrame(zeilen)

    try:
        import altair as alt
        diagramm = (
            alt.Chart(df)
            .mark_line(interpolate="step-after", point=True, strokeWidth=2)
            .encode(
                x=alt.X("stufe:Q",
                        scale=alt.Scale(domain=[0, 3], nice=False),
                        axis=alt.Axis(values=[0, 1, 2, 3],
                                      title="Aufwandsstufe k")),
                y=alt.Y("aehnlichkeit:Q",
                        scale=alt.Scale(domain=[0, 1]),
                        axis=alt.Axis(format="%",
                                      title="erreichbare Ähnlichkeit")),
                color=alt.Color("typ:N", title="Typ"),
                tooltip=[alt.Tooltip("typ:N", title="Typ"),
                         alt.Tooltip("stufe:Q", title="Stufe k"),
                         alt.Tooltip("aehnlichkeit:Q", title="Ähnlichkeit",
                                     format=".0%")],
            )
            .properties(height=320)
        )
        try:
            st.altair_chart(diagramm, width="stretch")
        except TypeError:   # aeltere Streamlit-Version
            st.altair_chart(diagramm, use_container_width=True)
    except Exception:
        # Rueckfall ohne Altair: einfache Liniendarstellung (ohne Treppenform).
        st.line_chart(df.pivot(index="stufe", columns="typ",
                               values="aehnlichkeit"))

    st.caption("Stufe 0 ist die ohne jede Änderung erreichte Übereinstimmung. "
               "Jede weitere Stufe schliesst die vorhergehenden ein, die Verläufe "
               "können daher nicht fallen.")


def seite_zusammenfassung(state, features, types, weights):
    st.title("Konfigurator — Typvergleich")

    if not state["units"]:
        st.info("Keine Betrachtungseinheiten vorhanden.")
        return

    stufen = [(core.AUFWAND_KEIN,   "k = 0  ohne Änderung"),
              (core.AUFWAND_GERING, "k = 1  nur geringe"),
              (core.AUFWAND_MITTEL, "k = 2  geringe + mittlere"),
              (core.AUFWAND_HOCH,   "k = 3  alle Änderungen")]

    st.divider()
    for uid in state["units"]:
        auswahl = core.get_engere_auswahl(state, uid)
        st.header(f"Betrachtungseinheit {uid}")
        if not auswahl:
            st.caption("Keine Typen in der engeren Auswahl.")
            st.divider()
            continue

        final = core.get_zieltyp(state, uid)

        # Gemeinsames Stufendiagramm ueber alle Typen der engeren Auswahl.
        _stufendiagramm(state, features, types, weights, uid, auswahl)

        # Kennzahlen und Zieltyp-Festlegung je Typ. Reihen zu hoechstens drei
        # Spalten, damit die Darstellung auch bei vielen Typen lesbar bleibt.
        pro_reihe = 3
        for start in range(0, len(auswahl), pro_reihe):
            gruppe = auswahl[start:start + pro_reihe]
            spalten = st.columns(pro_reihe)
            for sp, typ in zip(spalten, gruppe):
                with sp:
                    profil = types[typ]
                    score = core.soll_score_gestaffelt(state, uid, typ, profil,
                                                       features, weights)
                    vert = core.aufwand_verteilung(state, uid, typ, profil,
                                                   features)
                    titel = f"⭐ {typ}" if typ == final else typ
                    st.markdown(f"### {titel}")
                    for stufe, name in stufen:
                        wert = score[stufe]
                        st.markdown(f"{name}: **{wert * 100:.0f}%**")
                        st.progress(wert)
                    st.caption(
                        f"davon Änderungen: {vert[core.AUFWAND_GERING]} gering · "
                        f"{vert[core.AUFWAND_MITTEL]} mittel · "
                        f"{vert[core.AUFWAND_HOCH]} hoch")
                    # Finalen Zieltyp festlegen oder Festlegung aufheben.
                    if typ == final:
                        if st.button("★ Festlegung aufheben",
                                     key=f"unset_ziel_{uid}_{typ}",
                                     width="stretch"):
                            core.set_zieltyp(state, uid, None)
                            st.rerun()
                    else:
                        if st.button("Als Zieltyp festlegen",
                                     key=f"set_ziel_{uid}_{typ}", type="primary",
                                     width="stretch"):
                            core.set_zieltyp(state, uid, typ)
                            st.rerun()
        st.divider()




def _fmt_auspr(wert):
    """Anzeige einer Auspraegung; ein unbekanntes Ist erscheint als Gedankenstrich."""
    return "—" if wert is core.OFFEN else str(wert)


def _handlungsfeld_tabelle(felder):
    """Automatisch erzeugte Uebersicht der Handlungsfelder."""
    import pandas as pd
    zeilen = []
    for f in felder:
        zeilen.append({
            "Merkmal": f["merkmal"],
            "Ist": _fmt_auspr(f["ist"]),
            "Soll": f["soll"],
            "Aufwand": core.AUFWAND_LABEL.get(f["aufwand"], "—"),
            "Gewicht": f"{f['gewicht']:g}",
            "Ähnlichkeitsgewinn": f"+{f['gewinn'] * 100:.1f} %-Pkt.",
        })
    st.dataframe(pd.DataFrame(zeilen), hide_index=True, width="stretch")


def _portfolio_matrix(felder):
    """Aufwand-Wirkung-Portfolio: Aufwand auf der Abszisse, Ähnlichkeitsgewinn auf
    der Ordinate. Handlungsfelder mit hohem Gewinn und geringem Aufwand liegen links
    oben und sind die naheliegenden ersten Schritte. Der Gewinn ist die normierte
    Form des Gewichts und daher die anschaulichere Achse (Prozentpunkte)."""
    import pandas as pd
    if not felder:
        return
    # Handlungsfelder mit gleichem Aufwand UND gleichem Gewinn liegen im Portfolio
    # deckungsgleich. Sie werden daher zu einem Punkt zusammengefasst und gemeinsam
    # beschriftet, damit sich die Beschriftungen nicht ueberdecken.
    gruppen = {}
    for f in felder:
        schluessel = (f["aufwand"] or core.AUFWAND_HOCH, round(f["gewinn"], 6))
        gruppen.setdefault(schluessel, []).append(f["merkmal"])
    zeilen = [{"aufwand": a, "gewinn": g * 100,
               "merkmal": ", ".join(namen), "anzahl": len(namen)}
              for (a, g), namen in gruppen.items()]
    df = pd.DataFrame(zeilen)
    try:
        import altair as alt
        # Gemeinsame Achsen fuer beide Ebenen. Die Groesse (size) wird NUR auf die
        # Kreise angewendet, nicht auf den Text, da size bei mark_text sonst als
        # Schriftgroesse interpretiert wuerde und die Beschriftung riesig geraet.
        basis = alt.Chart(df).encode(
            x=alt.X("aufwand:Q", scale=alt.Scale(domain=[0.5, 3.5], nice=False),
                    axis=alt.Axis(values=[1, 2, 3], title="Aufwand",
                                  labelExpr="datum.value == 1 ? 'gering' : "
                                            "datum.value == 2 ? 'mittel' : 'hoch'")),
            y=alt.Y("gewinn:Q", scale=alt.Scale(zero=True),
                    axis=alt.Axis(title="Ähnlichkeitsgewinn (Prozentpunkte)")),
        )
        punkte = basis.mark_circle(opacity=0.7).encode(
            size=alt.Size("anzahl:Q", legend=None,
                          scale=alt.Scale(domain=[1, 6], range=[160, 420])),
            tooltip=[alt.Tooltip("merkmal:N", title="Merkmale"),
                     alt.Tooltip("aufwand:Q", title="Aufwand"),
                     alt.Tooltip("gewinn:Q", title="Ähnlichkeitsgewinn (%-Pkt.)",
                                 format=".1f")],
        )
        beschriftung = basis.mark_text(
            align="left", dx=12, fontSize=11, baseline="middle").encode(
            text="merkmal:N")
        diagramm = (punkte + beschriftung).properties(height=340)
        try:
            st.altair_chart(diagramm, width="stretch")
        except TypeError:
            st.altair_chart(diagramm, use_container_width=True)
    except Exception:
        st.dataframe(df, hide_index=True, width="stretch")
    st.caption("Handlungsfelder mit hohem Ähnlichkeitsgewinn und geringem Aufwand "
               "sind die naheliegenden ersten Schritte. Merkmale mit gleichem "
               "Aufwand und Gewicht sind zu einem Punkt zusammengefasst.")


def _massnahmen_eingabe(state, uid, felder):
    """Je Handlungsfeld die drei Eingaben des Anwenders: konkrete Massnahme,
    Etappe und Verantwortlichkeit. Das Werkzeug gibt inhaltlich nichts vor."""
    for f in felder:
        m = f["merkmal"]
        eintrag = core.get_massnahme(state, uid, m)
        with st.expander(
                f"{m}:  {_fmt_auspr(f['ist'])} → {f['soll']}"
                f"   ({core.AUFWAND_LABEL.get(f['aufwand'], '—')} Aufwand)"):
            text = st.text_area(
                "Massnahme", value=eintrag["text"], key=f"mn_txt_{uid}_{m}",
                placeholder="Was ist konkret zu tun, um die Soll-Auspraegung zu erreichen?",
                height=80)
            if text != eintrag["text"]:
                core.set_massnahme(state, uid, m, text=text)
            sp_e, sp_w = st.columns(2)
            with sp_e:
                stufen = list(core.AUFWAND_STUFEN)
                aktuell = eintrag["etappe"] or core.AUFWAND_HOCH
                wahl = st.selectbox(
                    "Etappe", options=stufen, index=stufen.index(aktuell),
                    format_func=lambda k: f"{k}. Etappe ({core.AUFWAND_LABEL[k]})",
                    key=f"mn_et_{uid}_{m}",
                    help="Aus dem Aufwand vorbelegt. Aenderbar, wenn Abhaengigkeiten "
                         "zwischen Massnahmen eine andere Reihenfolge erfordern.")
                if wahl != eintrag["etappe"]:
                    core.set_massnahme(state, uid, m, etappe=wahl)
            with sp_w:
                wer = st.text_input("Verantwortlich", value=eintrag["wer"],
                                    key=f"mn_wer_{uid}_{m}")
                if wer != eintrag["wer"]:
                    core.set_massnahme(state, uid, m, wer=wer)


def _morphologie_html(state, uid, typ_name, typ_profil, features):
    """Morphologischer Kasten fuer den Export: alle Merkmale mit ihren
    Auspraegungen, das Profil des Zieltyps grau hinterlegt.

    Die Merkmale sind nach der Erfassungsregel in zwei Bloecke geteilt. Bei
    bekanntem Ist-Zustand werden die Ist- und die Soll-Auspraegung gekennzeichnet,
    bei Uebereinstimmung beider als Ist = Soll. Bei unbekanntem Ist-Zustand wird
    allein die Soll-Auspraegung ausgewiesen."""
    bekannt, unbekannt = [], []
    for m in features:
        (bekannt if core.get_choice(state, uid, m) is not core.OFFEN
         else unbekannt).append(m)
    spalten = max((len(features[m]) for m in features), default=0)

    def zeile(m):
        optionen = list(features[m])
        profil = typ_profil.get(m, set())
        ist = core.get_choice(state, uid, m)
        soll = core.get_soll(state, uid, typ_name, m)
        if soll is core.OFFEN or soll == core.NICHTS_ANSTREBEN:
            soll = None
        i_ist = optionen.index(ist) if ist in optionen else None
        i_soll = optionen.index(soll) if soll in optionen else None

        zellen = []
        for j, a in enumerate(optionen):
            marke, klasse = "", "kp" if a in profil else ""
            if i_ist is not None and i_soll is not None and i_ist == i_soll:
                if j == i_ist:
                    marke = "Ist = Soll"
            elif j == i_ist:
                marke = "Ist"
            elif j == i_soll:
                marke = "Soll"
            stil = f' class="{klasse}"' if klasse else ""
            inhalt = f"{a}<br><span class='mk'>{marke}</span>" if marke else a
            zellen.append(f"<td{stil}>{inhalt}</td>")
        zellen += ["<td class='leer'></td>"] * (spalten - len(optionen))
        return f"<tr><th class='mn'>{m}</th>" + "".join(zellen) + "</tr>"

    teile = ["<table class='morph'>"]
    for titel, gruppe in (("Merkmale mit bekanntem Ist-Zustand", bekannt),
                          ("Merkmale mit unbekanntem Ist-Zustand", unbekannt)):
        if not gruppe:
            continue
        teile.append(f"<tr><th class='block' colspan='{spalten + 1}'>{titel}</th></tr>")
        teile.extend(zeile(m) for m in gruppe)
    teile.append("</table>")
    teile.append("<p class='leg'>Grau hinterlegt: Konfigurationsprofil des Zieltyps. "
                 "Bei bekanntem Ist-Zustand sind die Ist- und die Soll-Auspraegung "
                 "gekennzeichnet, bei Uebereinstimmung beider als Ist = Soll.</p>")
    return "\n".join(teile)


def _export_html(state, features, types, weights):
    """Erzeugt das Gesamtergebnis als HTML. Der Anwender kann daraus ueber die
    Druckfunktion seines Browsers ein PDF erstellen."""
    teile = ["<html><head><meta charset='utf-8'><title>Konfigurationsergebnis</title>",
             "<style>body{font-family:sans-serif;margin:2cm;line-height:1.5}",
             "table{border-collapse:collapse;width:100%;margin:1em 0}",
             "th,td{border:1px solid #999;padding:6px;text-align:left;font-size:10pt}",
             "th{background:#eee}h1{font-size:18pt}h2{font-size:14pt;margin-top:2em}",
             "h3{font-size:12pt}",
             "table.morph td{text-align:center;font-size:8.5pt;vertical-align:top}",
             "table.morph td.kp{background:#c8c8c8}",
             "table.morph td.leer{background:#fafafa;border:none}",
             "table.morph th.mn{width:18%;font-weight:600;background:#f4f4f4}",
             "table.morph th.block{background:#ddd;text-align:left;font-size:10pt}",
             "span.mk{font-weight:700;font-size:8pt;white-space:nowrap}",
             "p.leg{font-size:9pt;color:#444}</style></head><body>",
             "<h1>Ergebnis des Konfigurationsvorgehens</h1>"]
    for uid in state["units"]:
        typ = core.get_zieltyp(state, uid)
        teile.append(f"<h2>Betrachtungseinheit {uid}</h2>")
        if not typ:
            teile.append("<p>Kein Zieltyp festgelegt.</p>")
            continue
        profil = types[typ]
        score = core.soll_score_gestaffelt(state, uid, typ, profil, features, weights)
        teile.append(f"<p><b>Zieltyp:</b> {typ}</p>")
        teile.append("<p><b>Erreichbare Ähnlichkeit je Aufwandsstufe:</b> "
                     + " &nbsp; ".join(f"k={k}: {score[k]*100:.0f}%"
                                       for k in core.AUFWAND_STUFEN_K) + "</p>")
        felder = core.handlungsfelder(state, uid, typ, profil, features, weights)
        gruppen = core.massnahmen_nach_etappe(state, uid, felder)
        teile.append("<h3>Massnahmenplan</h3>")
        if not felder:
            teile.append("<p>Keine Handlungsfelder, der Zieltyp ist bereits erreicht.</p>")
        for etappe in core.AUFWAND_STUFEN:
            if not gruppen.get(etappe):
                continue
            teile.append(f"<h4>{etappe}. Etappe ({core.AUFWAND_LABEL[etappe]}er Aufwand)</h4>")
            teile.append("<table><tr><th>Merkmal</th><th>Ist</th><th>Soll</th>"
                         "<th>Ähnlichkeitsgewinn</th>"
                         "<th>Massnahme</th><th>Verantwortlich</th></tr>")
            for f in gruppen[etappe]:
                mn = core.get_massnahme(state, uid, f["merkmal"])
                teile.append(
                    f"<tr><td>{f['merkmal']}</td><td>{_fmt_auspr(f['ist'])}</td>"
                    f"<td>{f['soll']}</td>"
                    f"<td>+{f['gewinn'] * 100:.1f} %-Pkt.</td>"
                    f"<td>{mn['text'] or '—'}</td>"
                    f"<td>{mn['wer'] or '—'}</td></tr>")
            teile.append("</table>")
        teile.append("<h3>Morphologischer Kasten mit Zielkonfiguration</h3>")
        teile.append(_morphologie_html(state, uid, typ, profil, features))
        nicht = core.nicht_erreichte_merkmale(state, uid, typ, profil, features)
        if nicht:
            teile.append("<h3>Nicht erreichte Merkmale</h3><table>"
                         "<tr><th>Merkmal</th><th>Begruendung</th></tr>")
            for m, grund in nicht:
                teile.append(f"<tr><td>{m}</td><td>{grund}</td></tr>")
            teile.append("</table>")
    teile.append("</body></html>")
    return "\n".join(teile)


def seite_massnahmen(state, features, types, weights):
    st.title("Konfigurator — Massnahmenplanung")

    if not state["units"]:
        st.info("Keine Betrachtungseinheiten vorhanden.")
        return

    st.caption("Aus dem Vergleich von Ist- und Soll-Auspraegung des Zieltyps ergeben "
               "sich die Handlungsfelder. Das Werkzeug strukturiert sie, die "
               "inhaltliche Ausgestaltung der Massnahmen erfolgt durch den Anwender.")
    st.divider()

    for uid in state["units"]:
        typ = core.get_zieltyp(state, uid)
        st.header(f"Betrachtungseinheit {uid}")
        if not typ:
            st.caption("Kein Zieltyp festgelegt — uebersprungen.")
            st.divider()
            continue
        st.caption(f"Zieltyp: **{typ}**")

        profil = types[typ]
        felder = core.handlungsfelder(state, uid, typ, profil, features, weights)
        core.etappe_vorbelegen(state, uid, felder)

        if not felder:
            st.success("Keine Handlungsfelder: Der Zieltyp ist im aktuellen Zustand "
                       "bereits erreicht.")
        else:
            st.subheader("Handlungsfelder")
            _handlungsfeld_tabelle(felder)

            st.subheader("Priorisierung")
            _portfolio_matrix(felder)

            st.subheader("Massnahmen")
            _massnahmen_eingabe(state, uid, felder)

        nicht = core.nicht_erreichte_merkmale(state, uid, typ, profil, features)
        if nicht:
            with st.expander(f"Nicht erreichte Merkmale ({len(nicht)})"):
                st.caption("Diese Merkmale begruenden, warum die erreichbare "
                           "Ähnlichkeit unter 100 % bleibt.")
                for m, grund in nicht:
                    st.markdown(f"- **{m}**: {grund}")
        st.divider()

    st.subheader("Ergebnis exportieren")
    st.download_button(
        "Gesamtergebnis als HTML herunterladen",
        data=_export_html(state, features, types, weights),
        file_name="konfigurationsergebnis.html", mime="text/html")
    st.caption("Die Datei laesst sich im Browser oeffnen und ueber dessen "
               "Druckfunktion als PDF speichern.")


init_state()
state = st.session_state.state

# --- Morphologie beschaffen: in Schritt 1 via Hochladefeld, danach aus dem Cache ---
# Das Hochladefeld erscheint nur in der Zustandserfassung. Sobald eine gueltige
# Morphologie geladen ist, wird sie in st.session_state.morphologie gehalten und
# in den folgenden Schritten von dort genutzt (ohne erneutes Hochladefeld).
features = types = weights = None
quelle_ok = False
gewicht_problem = None
morph_cache = st.session_state.get("morphologie")

if state["phase"] == "konfiguration":
    st.sidebar.header("Morphologie")
    hochgeladen = st.sidebar.file_uploader(
        "Excel-Morphologie laden (.xlsx)", type=["xlsx"],
        help="Datei mit Merkmalen, Auspraegungen, Typprofilen und (optional) Gewichten."
    )
    try:
        if hochgeladen is not None:
            features, types, weights = lade_aus_bytes(hochgeladen.getvalue())
            st.session_state.morphologie = (features, types, weights)
            st.sidebar.caption(f"Geladen: {hochgeladen.name}")
            quelle_ok = True
        elif morph_cache is not None:
            features, types, weights = morph_cache
            quelle_ok = True
        else:
            st.info("Bitte links eine Excel-Morphologie hochladen, um zu starten.")
    except core.GewichtFehler as e:
        # Gewichte fehlerhaft. Morphologie ist trotzdem da (im Fehler mitgegeben).
        gewicht_problem = (e.probleme, e.features, e.types)

    # Gewichtsfehler behandeln (Wahl anbieten).
    if gewicht_problem is not None:
        probleme, feat_roh, typ_roh = gewicht_problem
        st.error("Die Gewichtsspalte enthaelt fehlerhafte Werte:")
        for merkmal, grund in probleme:
            st.write(f"- **{merkmal}**: {grund}")
        st.write("Du kannst die Excel korrigieren und neu laden, "
                 "oder ohne Gewichtung fortfahren (alle Merkmale gleich).")
        if st.button("Ohne Gewichtung fortfahren", type="primary"):
            st.session_state.ignoriere_gewichte = True
            st.rerun()
        if st.session_state.get("ignoriere_gewichte"):
            features, types, weights = feat_roh, typ_roh, None
            st.session_state.morphologie = (features, types, weights)
            quelle_ok = True
            gewicht_problem = None
            st.info("Es wird ohne Gewichtung gerechnet (alle Merkmale gleich).")
else:
    # Spaetere Schritte: Morphologie aus dem Cache, kein Hochladefeld noetig.
    if morph_cache is not None:
        features, types, weights = morph_cache
        quelle_ok = True

# Ohne gueltige Quelle hier stoppen.
if not quelle_ok:
    st.stop()

# Gewichte einmalig initialisieren (aus Excel-Startwerten, falls vorhanden, sonst
# 1) und ab hier als editierbare Quelle verwenden. Die Zustandserfassung schreibt
# in state["gewichte"]; alle Berechnungen lesen von hier.
core.gewicht_init(state, features, weights)
weights = state["gewichte"]

# Aufgeschobenes Loeschen zuerst.
if "zu_loeschen" in st.session_state:
    core.remove_unit(state, st.session_state.zu_loeschen)
    del st.session_state.zu_loeschen
    st.session_state.pop("warte_bestaetigung", None)

# Seitenleiste: phasenabhaengige Navigation und Hinweise.
sidebar_navigation(state, features, types)

# Hauptbereich je nach Phase.
if state["phase"] == "konfiguration":
    seite_erfassung(state, features, weights)
elif state["phase"] == "ergebnis_soll":
    seite_ergebnis_soll(state, features, types, weights)
elif state["phase"] == "detaillierung":
    seite_detaillierung(state, features, types, weights)
elif state["phase"] == "zusammenfassung":
    seite_zusammenfassung(state, features, types, weights)
elif state["phase"] == "massnahmen":
    seite_massnahmen(state, features, types, weights)
else:
    # Fallback: alte Ist-Ergebnis-Seite (wird im neuen Fluss nicht mehr angesteuert)
    seite_ergebnis_soll(state, features, types, weights)