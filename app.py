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
        if core.get_name(state, uid) != uid:
            spalte.caption(core.get_name(state, uid))

        if st.session_state.get("warte_bestaetigung") == uid:
            spalte.caption("⚠ löschen?")
            if spalte.button("Löschen", key=f"ok_{uid}_{merkmal}", type="tertiary",
                             help=f"{core.get_name(state, uid)} endgültig löschen", width="stretch"):
                st.session_state.zu_loeschen = uid
                st.rerun()
            if spalte.button("Abbrechen", key=f"no_{uid}_{merkmal}", type="tertiary",
                             help="Löschen abbrechen", width="stretch"):
                st.session_state.pop("warte_bestaetigung", None)
                st.rerun()
        else:
            if spalte.button("🗑", key=f"del_{uid}_{merkmal}", help=f"{core.get_name(state, uid)} löschen"):
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


@st.dialog("Betrachtungseinheit löschen")
def _loeschen_dialog(state, uid):
    """Sicherheitsabfrage vor dem Loeschen einer Betrachtungseinheit. Als Fenster,
    damit der Tabellenkopf uebersichtlich bleibt."""
    st.markdown(f"**{core.get_name(state, uid)}** wirklich löschen?")
    st.caption("Alle Angaben dieser Betrachtungseinheit gehen dabei verloren. "
               "Der Schritt lässt sich nicht rückgängig machen.")
    sp1, sp2 = st.columns(2)
    if sp1.button("Löschen", type="primary", width="stretch",
                  key=f"dlg_del_ok_{uid}"):
        st.session_state.zu_loeschen = uid
        st.rerun()
    if sp2.button("Abbrechen", width="stretch", key=f"dlg_del_no_{uid}"):
        st.rerun()


@st.dialog("Betrachtungseinheit umbenennen")
def _umbenennen_dialog(state, uid):
    """Modaler Dialog zum Umbenennen einer Einheit. Zeigt ein Textfeld mit dem
    aktuellen Namen und speichert erst auf Bestaetigung. Die interne unit_id
    (Buchstabe) bleibt unveraendert; nur der Anzeigename wird gesetzt."""
    aktuell = state.get("namen", {}).get(uid, "")
    neuer = st.text_input("Name der Betrachtungseinheit", value=aktuell,
                          placeholder="z. B. Werk Nord")
    links, rechts = st.columns(2)
    if links.button("Speichern", type="primary", width="stretch"):
        core.set_name(state, uid, neuer)
        st.rerun()
    if rechts.button("Abbrechen", width="stretch"):
        st.rerun()


def erfassung_tabelle(state, features, merkmal, weights):
    """Gemeinsame Erfassungstabelle fuer EIN Merkmal unter Tims Regel.
    Pro Einheit ein Welt-Dropdown (Ist-bekannt / Ist-unbekannt) und darunter
    je Auspraegung EIN Knopf, dessen Wirkung von der Welt abhaengt (core.klick_erfassung):
      - Ist-bekannt: erster Klick setzt Ist, weitere Klicks auf andere Auspraegungen
        setzen Potenzial; Klick auf das Ist setzt das Merkmal zurueck.
      - Ist-unbekannt: Klick schaltet Ausschluss an/aus (beliebig viele)."""
    optionen = features[merkmal]

    # Textspalte + je Einheit eine Spalte + Anlege-Spalte.
    breiten = [4] + [2] * len(state["units"]) + [1]

    # Kopfzeile: Einheitenname + Loeschen; Anlegen rechts.
    kopf = st.columns(breiten)
    kopf[0].markdown("")
    for i, uid in enumerate(state["units"]):
        spalte = kopf[1 + i]
        titel, k_ren, k_del = spalte.columns([2, 1, 1])
        titel.markdown(f"**{core.get_name(state, uid)}**")
        if k_ren.button("✏", key=f"ren_{uid}_{merkmal}",
                        help=f"{core.get_name(state, uid)} umbenennen"):
            _umbenennen_dialog(state, uid)
        if k_del.button("🗑", key=f"del_{uid}_{merkmal}",
                        help=f"{core.get_name(state, uid)} löschen"):
            _loeschen_dialog(state, uid)

    if kopf[-1].button("➕", key=f"add_{merkmal}", help="Neue Betrachtungseinheit",
                       width="stretch"):
        core.add_unit(state, features)
        st.rerun()

    # Welt-Dropdown je Einheit.
    weltzeile = st.columns(breiten)
    weltzeile[0].caption("Ausprägungen:")
    welt_labels = {core.WELT_IST_BEKANNT: "Ist bekannt",
                   core.WELT_IST_UNBEKANNT: "Ist unbekannt"}
    for i, uid in enumerate(state["units"]):
        akt = core.get_welt(state, uid, merkmal)
        wahl = weltzeile[1 + i].selectbox(
            "Bekannt?",
            options=[core.WELT_IST_BEKANNT, core.WELT_IST_UNBEKANNT],
            index=0 if akt == core.WELT_IST_BEKANNT else 1,
            format_func=lambda w: welt_labels[w],
            key=f"welt_{uid}_{merkmal}",
            help="Ist für dieses Merkmal die aktuelle Ausprägung bekannt?",
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
            # Der Hinweis richtet sich danach, was der Klick bewirken wuerde.
            if welt == core.WELT_IST_UNBEKANNT:
                hilfe = ("Ausschluss (Klick setzt Ausschluss zurück)"
                         if zustand == core.ZELLE_AUSSCHLUSS
                         else "Als Ausschluss markieren")
            elif core.get_choice(state, uid, merkmal) is core.OFFEN:
                hilfe = "Als Ist markieren"
            elif zustand == core.ZELLE_IST:
                hilfe = "Ist (Klick setzt Merkmal zurück)"
            elif zustand == core.ZELLE_POTENTIAL:
                hilfe = "Potenzial (Klick setzt Potenzial zurück)"
            else:
                hilfe = "Als Potenzial markieren"
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


@st.dialog("Was passiert auf dieser Seite?", width="large")
def _hilfe_dialog(text):
    """Erklaerung zur aktuellen Seite als Fenster. Aufbau je Seite gleich:
    Einordnung ins Vorgehen, Ergebnis der Seite, Bedienschritte mit Begriffs-
    erklaerungen, Hinweis zum Weitergehen.

    Absaetze, die mit '>>' beginnen, gehoeren zu einem Unterpunkt und werden
    eingerueckt. Der Einzug entsteht ueber eine schmale Randspalte und wirkt
    dadurch auf den ganzen Absatz, also auch auf umgebrochene Zeilen. Eine
    Einrueckung ueber fuehrende Leerzeichen wuerde nur die erste Zeile treffen,
    eine ueber ein HTML-Element wuerde die Fett- und Kursivauszeichnung
    unwirksam machen."""
    absaetze = [" ".join(a.split()) for a in text.strip().split("\n\n")]
    i = 0
    while i < len(absaetze):
        if absaetze[i].startswith(">>"):
            # Aufeinanderfolgende eingerueckte Absaetze in einem Block setzen,
            # damit die Abstaende gleichmaessig bleiben.
            block = []
            while i < len(absaetze) and absaetze[i].startswith(">>"):
                block.append(absaetze[i][2:].strip())
                i += 1
            rand, inhalt = st.columns([1, 12])
            with inhalt:
                # HTML erlaubt, damit farbige Legendenzeichen die Diagrammfarben
                # aufgreifen koennen.
                st.markdown("\n\n".join(block), unsafe_allow_html=True)
        else:
            st.markdown(absaetze[i], unsafe_allow_html=True)
            i += 1


def _hilfe_text(phase):
    """Erklaertext zur Phase, oder None, wenn fuer sie keiner hinterlegt ist.
    Die Zuordnung steht in einer Funktion, damit sie erst beim Aufruf aufgeloest
    wird und die Texte weiter bei ihrer jeweiligen Seite stehen koennen."""
    return {
        "konfiguration": HILFE_ERFASSUNG,
        "ergebnis_soll": HILFE_AEHNLICHKEIT,
        "detaillierung": HILFE_ZIELTYP,
        "zusammenfassung": HILFE_TYPVERGLEICH,
        "massnahmen": HILFE_KONSOLIDIERUNG,
    }.get(phase)


def _sidebar_hilfe(phase):
    """Hilfe-Knopf oben in der Seitenleiste. Das Fenster oeffnet sich
    ausschliesslich auf Knopfdruck."""
    text = _hilfe_text(phase)
    if text is None:
        return
    with st.sidebar:
        # Eigene Einfaerbung, damit sich der Hilfe-Knopf klar von den
        # hervorgehobenen Navigationsknoepfen (Weiter/Zurueck) unterscheidet.
        st.markdown(
            f"""<style>
            div.st-key-hilfe_{phase} button {{
                background-color: #1f77b4; border-color: #1f77b4; color: #ffffff;
            }}
            div.st-key-hilfe_{phase} button:hover {{
                background-color: #17608f; border-color: #17608f; color: #ffffff;
            }}
            </style>""", unsafe_allow_html=True)
        geklickt = st.button("ℹ️  Was passiert auf dieser Seite?",
                             key=f"hilfe_{phase}", width="stretch")
        st.divider()
    if geklickt:
        _hilfe_dialog(text)


HILFE_ERFASSUNG = """
**Schritt 1 von 4 · Zustandsaufnahme**

Hier beschreiben Sie den heutigen Zustand Ihrer Auftragsabwicklung und den
Spielraum für künftige Veränderungen.

**1. Legen Sie über ➕ eine erste Betrachtungseinheit an.**

*Betrachtungseinheit: ein Bereich Ihres Unternehmens, dessen Auftragsabwicklung
sich einheitlich beschreiben lässt. Sinnvoll sind Bereiche, die es real gibt und
die sich in ihrer Auftragsabwicklung erkennbar unterscheiden. Über ✏ benennen Sie
eine Betrachtungseinheit um, über 🗑 löschen Sie sie.*

*Bsp: eine Serienfertigung von Antriebskomponenten als erste
Betrachtungseinheit und eine Sondermaschinenfertigung als zweite
Betrachtungseinheit.*

**2. Entscheiden Sie je Merkmal zuerst, ob Sie die aktuelle Ausprägung kennen.**

>>**2.1 Ausprägung bekannt.** Wählen Sie im Auswahlfeld *Ist
bekannt* und klicken Sie die zutreffende Ausprägung an. Ein Klick auf eine weitere
Ausprägung setzt zusätzlich ein Potenzial.

>>*Ist-Ausprägung ✗: die aktuell zutreffende Ausprägung. Je
Merkmal und Betrachtungseinheit ist genau eine möglich.*

>>*Potenzial-Ausprägung 🟠: eine Ausprägung, die aktuell
nicht zutrifft, für diese Betrachtungseinheit aber erreichbar wäre. Mehrere sind
möglich, die Angabe ist freiwillig.*

>>*Bsp: Ein Unternehmen verkauft seine Produkte heute, könnte
sie künftig aber auch vermieten. Verkauf ist dann das Ist, Vermietung ein
Potenzial.*

>>Treffen mehrere Ausprägungen zugleich zu, ist die
Betrachtungseinheit zu grob abgegrenzt. Legen Sie dann über ➕ eine weitere an,
sofern sie einem real abgrenzbaren Bereich entspricht, andernfalls wählen Sie die
überwiegend zutreffende Ausprägung.

>>**2.2 Ausprägung unbekannt.** Wählen Sie im Auswahlfeld
*Ist unbekannt*. Die Angabe von Ist- und Potenzial-Ausprägungen entfällt dadurch,
Sie können jedoch Ausprägungen ausschließen.

>>*Ausschluss ⛔: eine Ausprägung, die für diese
Betrachtungseinheit nicht in Frage kommt. Mehrere sind möglich, und die Angabe ist
freiwillig. Was Sie nicht ausschließen, bleibt als Spielraum erhalten.*

>>*Bsp: Ein Hersteller stoffschlüssig gefügter Produkte weiß
noch nicht, wie er eine Demontage gestalten würde, kann aber eine zerstörungsfreie
Demontage ausschließen.*

**3. Wiederholen Sie das für alle Merkmale und alle Betrachtungseinheiten.**
Merkmale dürfen offen bleiben, wenn Sie die Ausprägung nicht kennen und nichts
ausschließen möchten.

**4. Gewichten Sie die Merkmale nach ihrer Bedeutung.** Das Feld steht rechts neben
jedem Merkmalsnamen.

*Gewicht: die relative Wichtigkeit eines Merkmals für die spätere Ähnlichkeit. 1
bedeutet normal, 2 doppelt so wichtig.*

Sobald alle Betrachtungseinheiten erfasst sind, geht es links über den Knopf zur
Ähnlichkeitsbewertung weiter.
"""


def seite_erfassung(state, features, weights):
    """Schritt 1+3 vereint: gemeinsame Erfassung von Ist, Potenzial, Ausschluss."""
    st.title("Zustandsaufnahme")
    st.divider()

    for idx, merkmal in enumerate(features):
        sp_titel, sp_gewicht = st.columns([4, 1])
        with sp_titel:
            st.caption("Merkmal:")
            st.subheader(merkmal)
        with sp_gewicht:
            g = st.number_input(
                "Gewicht", min_value=1.0,
                value=float(core.get_gewicht(state, merkmal)),
                step=0.5, key=f"gw_{merkmal}",
                help="Relative Wichtigkeit für die Ähnlichkeit: 1 = normal, "
                     "2 = doppelt so wichtig, 1,5 = anderthalbfach usw. "
                     "Gilt für das Merkmal über alle Betrachtungseinheiten.")
            core.set_gewicht(state, merkmal, g)
        erfassung_tabelle(state, features, merkmal, weights)
        if idx < len(features) - 1:
            st.divider()


def sidebar_navigation(state, features, types):
    """Phasenabhaengige Seitenleiste: Navigation (Weiter/Zurueck) und Hinweis,
    was zum Weiterkommen noch offen ist. Das Hochladefeld und der volle
    Erfassungsstatus erscheinen nur in Schritt 1 (Zustandsaufnahme)."""
    phase = state["phase"]
    _sidebar_hilfe(phase)
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
        st.header("Nächster Schritt")
        if st.button("← Zurück zur Zustandsaufnahme", width="stretch"):
            state["phase"] = "konfiguration"
            st.rerun()
        offen = [core.get_name(state, uid) for uid in state["units"]
                 if not core.get_engere_auswahl(state, uid)]
        alle_haben = bool(state["units"]) and not offen
        if st.button("Weiter zur Zieltypbestimmung →", type="primary",
                     width="stretch", disabled=not alle_haben):
            state["phase"] = "detaillierung"
            st.rerun()
        st.divider()
        if alle_haben:
            st.caption("Für jede Betrachtungseinheit ist mindestens ein Typ in der engeren "
                       "Auswahl. Sie können zur Zieltypbestimmung weitergehen.")
        else:
            st.caption("Wählen Sie für jede Betrachtungseinheit mindestens einen "
                       "Typ in die engere Auswahl. Erst dann können Sie weiter. Es "
                       "fehlt noch: " + ", ".join(offen) + ".")


def _sidebar_sollfestlegung(state, features):
    """Schritt 3 (Zieltypbestimmung): Navigation + Hinweis, was je Fall zu tun ist."""
    with st.sidebar:
        st.header("Nächster Schritt")
        if st.button("← Zurück zur Ähnlichkeitsbewertung", width="stretch"):
            state["phase"] = "ergebnis_soll"
            st.rerun()
        # Belegt zugleich das Soll vor (Faelle 1/2/3) und prueft Vollstaendigkeit.
        vollstaendig = core.alle_detaillierungen_vollstaendig(state, features)
        if st.button("Weiter zum Typvergleich →", type="primary",
                     width="stretch", disabled=not vollstaendig):
            state["phase"] = "zusammenfassung"
            st.rerun()
        st.divider()
        if not vollstaendig:
            st.warning("Es fehlen noch Angaben. Erst wenn überall eine Angabe "
                       "vorliegt, geht es weiter.")


def _sidebar_typvergleich(state):
    """Schritt 4 (Typvergleich): Navigation + Hinweis zur Zieltyp-Festlegung."""
    with st.sidebar:
        st.header("Nächster Schritt")
        if st.button("← Zurück zur Zieltypbestimmung", width="stretch"):
            state["phase"] = "detaillierung"
            st.rerun()
        # Weiter erst, wenn jede Einheit mit engerer Auswahl einen Zieltyp hat.
        offen = [uid for uid in state["units"]
                 if core.get_engere_auswahl(state, uid)
                 and core.get_zieltyp(state, uid) is None]
        if st.button("Weiter zur Konsolidierung →", type="primary",
                     width="stretch", disabled=bool(offen)):
            state["phase"] = "massnahmen"
            st.rerun()
        st.divider()
        st.caption("Legen Sie je Betrachtungseinheit einen finalen Zieltyp fest.")
        if offen:
            st.warning("Noch kein Zieltyp festgelegt für: " + ", ".join(offen))


def _sidebar_massnahmen(state):
    """Schritt 4 (Konsolidierung): Zurueck-Navigation + Hinweis."""
    with st.sidebar:
        st.header("Konsolidierung")
        if st.button("← Zurück zum Typvergleich", width="stretch"):
            state["phase"] = "zusammenfassung"
            st.rerun()
        st.divider()
        st.caption("Die Änderungen ergeben sich aus dem Vergleich von Ist und "
                   "Ziel des Zieltyps. Führe sie hier unternehmensweit zusammen, "
                   "harmonisiere sie und priorisiere die Betrachtungseinheiten.")


def _sidebar_erfassung(state, features):
    """Seitenleiste: Stand je Einheit + Auswertungsknopf."""
    with st.sidebar:
        st.header("Fortschritt")

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
                        text=f"{core.get_name(state, uid)}: {versorgt}/{gesamt} "
                             "Merkmale bearbeitet"
                             + ("  ✓" if offen == 0 else f"  ({offen} offen)"))

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
                f"- {core.get_name(state, uid)}: {len(ms)} Merkmal(e) mit "
                f"'Ist bekannt' aber ohne Ist"
                for uid, ms in pro_einheit.items()
            )
            st.warning(
                "Bevor es weitergeht, fehlt noch etwas:\n\n" + zeilen
                + "\n\nSetze dort entweder ein Ist, oder stelle das Merkmal auf "
                "'Ist unbekannt', wenn Sie die aktuelle Ausprägung nicht kennen."
            )


def seite_konfiguration(state, features, weights):
    st.title("Ist-Aufnahme")

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
        core.STATUS_IST:           ("✅", "Passt aktuell"),
        core.STATUS_POTENTIAL:     ("🟠", "Passt potenziell"),
        core.STATUS_IST_UNPASSEND: ("🔶", "Passt nicht"),
        core.STATUS_OFFEN:         ("⬜", "Offen"),
        core.STATUS_BLOCKIERT:     ("⛔", "Ausgeschlossen"),
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
        if core.get_name(state, uid) != uid:
            kopf[i + 1].caption(core.get_name(state, uid))

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
    st.title("Potenzial-Aufnahme")

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
    Kein Ziel, keine harten Verstoesse (Potential ist hier noch nicht erfasst)."""
    import pandas as pd
    import altair as alt
    FARBE_IST = "#1f77b4"   # Blau (durchgaengig fuer Ist)

    st.title("Ist-Ergebnis")

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

        st.header(core.get_name(state, uid))

        # Kopf-Kennzahlen: nur Ist.
        k1, k2, k3 = st.columns(3)
        k1.metric("Bestpassender Typ", best_name)
        k2.metric("Ist-Übereinstimmung", _fmt(best_ist))
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
        st.subheader("Güte der Aussage")
        g1, g2 = st.columns(2)
        g1.metric("Abdeckung (beantwortete Merkmale)", _fmt(anteil),
                  help="Gewichteter Anteil der Merkmale, die nicht 'keine Angabe' sind.")
        g2.metric(f"Ist {best_name} (nur beantwortete)",
                  _fmt(kpi_beantwortet[best_name]),
                  help="Wie der Ist-Wert, aber nur über beantwortete Merkmale gerechnet.")
        if anteil < 1.0:
            st.caption("Diese Einheit hat 'keine Angabe'-Merkmale; der Ist-KPI ist "
                       "dadurch niedriger als der Wert über nur beantwortete Merkmale.")

        st.divider()


HILFE_AEHNLICHKEIT = """
**Schritt 2 von 4 · Ähnlichkeitsbewertung**

Das Werkzeug vergleicht Ihre Angaben mit jedem Auftragsabwicklungstyp und berechnet
die Ähnlichkeit automatisch. Ihre Aufgabe ist es, je Betrachtungseinheit die
aussichtsreichsten Typen in die engere Auswahl zu übernehmen.

*Auftragsabwicklungstyp: eine in der Realität typisch vorliegende Ausgestaltung der
Auftragsabwicklung. Für jeden Typ existiert ein Profil, das zeigt, welche
Ausprägungen je Merkmal für ihn charakteristisch sind. Je mehr Ihrer Angaben im
Profil eines Typs liegen, desto ähnlicher ist Ihre Betrachtungseinheit diesem Typ.*

**Aufbau der Seite**

**1. Tab je Betrachtungseinheit.** Jede Betrachtungseinheit wird für sich bewertet.

**2. Ähnlichkeit je Typ.** Oben stehen die bestpassenden Typen, darunter zeigt ein
Diagramm je Typ zwei Intervalle.

Waren alle Merkmale bekannt, ist die Ähnlichkeit eine einzelne Zahl. Sobald
Merkmale ohne Ist-Angabe geblieben sind, rechnet das Werkzeug zwei Annahmen durch:

*Maximum (rechtes Ende des Intervalls): der günstigste Fall, alle unbekannten
Merkmale passen zum Typ.*

*Minimum (linkes Ende des Intervalls): der ungünstigste Fall, alle unbekannten
Merkmale passen nicht zum Typ.*

Der tatsächliche, bisher unbekannte Wert liegt in diesem Intervall. Je breiter das
Intervall, desto unsicherer die Aussage.

*Bereinigter Wert (Strich im Intervall): Hier bleiben die Merkmale ohne Ist-Angabe
ganz außen vor, gerechnet wird nur über die bekannten Merkmale. Das ist die
belastbarste Schätzung und die Grundlage der Rangfolge.*

Je Typ werden zwei Intervalle gezeigt.

<span style="color:#1f77b4">■</span> *Ist: wie ähnlich die Betrachtungseinheit dem
Typ aktuell ist. Es zählen nur die Ist-Ausprägungen.*

<span style="color:#ff7f0e">■</span> *Potenzial: wie ähnlich sie dem Typ werden
könnte, wenn sie ihren Spielraum nutzt. Hier zählen zusätzlich die
Potenzial-Ausprägungen.*

**3. Gütebewertung.** Zwei Kennzahlen sagen, wie belastbar das Ergebnis ist.

*Abdeckung: gewichteter Anteil der bekannten Merkmale, also jener mit Ist-Angabe.
Je höher, desto schmaler die Intervalle und desto aussagekräftiger das Ergebnis.*

*Eindeutigkeit: Abstand zwischen bestem und zweitbestem Typ. Ein kleiner Abstand
heißt, dass mehrere Typen ähnlich gut passen.*

**4. Abgleich je Merkmal.** Klappen Sie einen Typ auf, um zu sehen, woher sein
Ähnlichkeitswert kommt. Je Merkmal steht dort einer von fünf Fällen.

*Merkmal bekannt*

>>*✅ Passt aktuell: Ihre aktuelle Ausprägung entspricht dem Profil des Typs.*

>>*🟠 Passt potenziell: passt aktuell nicht, aber eine angegebene
Potenzial-Ausprägung passt.*

>>*🔶 Passt nicht: Weder die aktuelle noch eine angegebene Potenzial-Ausprägung
passt.*

*Merkmal unbekannt*

>>*⬜ Offen: Das Profil dieses Typs wurde nicht vollständig ausgeschlossen, der Typ
bleibt hier also erreichbar. Ob die künftig angestrebte Ausprägung zu ihm passt,
ist noch offen.*

>>*⛔ Ausgeschlossen: Das Profil dieses Typs wurde vollständig ausgeschlossen, der
Typ ist hier also nicht erreichbar. Es liegt ein harter Verstoß vor.*

>>*Harter Verstoß: ein Merkmal, bei dem sämtliche für einen Typ passenden
Ausprägungen ausgeschlossen wurden. Der Typ ist damit nicht insgesamt
ausgeschlossen, es ist lediglich ein Warnhinweis, dass er bei diesem Merkmal
dauerhaft nicht passen wird.*

**5. Engere Auswahl.** Wählen Sie je Betrachtungseinheit die Typen, die Sie
weiterverfolgen möchten.

*Engere Auswahl: die Typen, für die Sie im nächsten Schritt eine Zielkonfiguration
ausarbeiten. Sinnvoll sind Typen mit hoher Ähnlichkeit, insbesondere im Potenzial.
Je mehr Sie wählen, desto aufwändiger wird der nächste Schritt.*

Sobald für jede Betrachtungseinheit mindestens ein Typ gewählt ist, geht es links
über den Knopf zur Zieltypbestimmung weiter.
"""


def seite_ergebnis_soll(state, features, types, weights):
    """Dashboard (Phase 'ergebnis_soll'): dreiteiliges Aehnlichkeitsmass je Typ
    als Intervall (min..max) mit bereinigtem Wert, fuer Ist (blau) und
    Potenzial (orange) uebereinander. Zahlen einklappbar, Guete-KPIs,
    Vergleichsfall je Typ und Typfestlegung."""
    import pandas as pd
    import altair as alt
    FARBE_IST = "#1f77b4"    # Blau
    FARBE_SOLL = "#ff7f0e"   # Orange

    st.title("Ähnlichkeitsbewertung")

    if not state["units"]:
        st.caption("Keine Betrachtungseinheiten vorhanden.")
        return

    _tabs = st.tabs([core.get_name(state, uid) for uid in state["units"]])
    for _tab, uid in zip(_tabs, state["units"]):
        with _tab:
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
            z1a.metric("Bestpassender Typ im Ist", best_ist_namen)
            z1b.metric("Ist bereinigt",
                       _fmt(best_ist_wert) if best_ist_wert is not None else "—")
            # Zeile 2: Potenzial
            z2a, z2b = st.columns([2, 1])
            z2a.metric("Bestpassender Typ im Potenzial", best_pot_namen)
            z2b.metric("Potenzial bereinigt",
                       _fmt(best_pot_wert) if best_pot_wert is not None else "—")

            # --- Intervall-Diagramm: pro Typ Ist (blau) und Soll (orange) ---
            st.subheader("Ähnlichkeit je Typ")

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
                    x=alt.X("min:Q", title="Ähnlichkeit (%)", scale=alt.Scale(domain=[0, 100])),
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
            st.subheader("Gütebewertung")
            g1, g2, g3 = st.columns(3)
            g1.metric("Abdeckung (Ist)", _fmt(anteil),
                      help="Gewichteter Anteil der bekannten Merkmale, also jener "
                           "mit Ist-Angabe. Je höher, desto schmaler die Intervalle "
                           "und desto aussagekräftiger das Ergebnis.")
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
                hinweise.append("Diese Betrachtungseinheit hat Merkmale ohne "
                                "Ist-Angabe, deshalb spannen die Ist-Balken einen "
                                "Bereich auf.")
            if hv_best > 0:
                hinweise.append(f"Achtung: {hv_best} Merkmal(e) können den Typ "
                                f"{best_name} dauerhaft nicht erreichen, weil alle "
                                "dafür passenden Ausprägungen ausgeschlossen sind.")
            for h in hinweise:
                st.caption(h)

            # --- Merkmals-Status je Typ (einklappbar, standardmaessig zu) ---
            st.subheader("Abgleich je Merkmal")
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
                            "Abgleich": f"{symbol} {klartext}",
                        })
                    st.dataframe(pd.DataFrame(zeilen_status), hide_index=True, width="stretch")

            # --- Engere Auswahl fuer die Detaillierung (bis zu 3 Typen) ---
            st.subheader("Engere Auswahl")
            typ_namen = [n for n, _ in rang_ist]
            aktuelle_auswahl = core.get_engere_auswahl(state, uid)
            wahl = st.multiselect(
                f"Typen für Betrachtungseinheit {uid} "
                "(werden anschließend ausgestaltet und verglichen)",
                options=typ_namen,
                default=[t for t in aktuelle_auswahl if t in typ_namen],
                key=f"engere_auswahl_{uid}",
                placeholder="Typen wählen",
                help="Je mehr Typen gewählt sind, desto aufwändiger wird die "
                     "Zieltypbestimmung, da dort jedes Merkmal je Typ einzeln zu "
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
                st.caption("Noch keine Typen gewählt. Wählen Sie die Typen, für die "
                           "Sie im nächsten Schritt eine Zielkonfiguration "
                           "ausarbeiten möchten.")
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


@st.dialog("Aufwand und Kosten angeben")
def _aufwand_dialog(state, uid, typ, merkmal, ziel_opt, im_profil):
    """Modal beim Waehlen einer Ziel-Auspraegung, die vom Ist abweicht.
    Der Aufwand ist Pflicht, die Kosten sind optional. Die Ziel-Auspraegung
    wird erst beim Speichern uebernommen; Abbrechen laesst alles unveraendert."""
    st.markdown(f"**{merkmal}**")
    st.markdown(f"Ziel-Ausprägung: **{ziel_opt}**")
    if not im_profil:
        st.info("Diese Ausprägung liegt außerhalb des Profils und erhöht die "
                "Ähnlichkeit zum Typ nicht. Der Aufwand wird dennoch erfasst.")
    stufen = list(core.AUFWAND_STUFEN)
    aktueller = core.get_aufwand(state, uid, typ, merkmal)
    vor = stufen.index(aktueller) if aktueller in stufen else None
    stufe = st.radio("Aufwand (Pflicht)", stufen,
                     format_func=lambda x: core.AUFWAND_LABEL[x],
                     index=vor, horizontal=True)
    kosten = st.number_input("Kosten in Euro (optional)", min_value=0.0,
                             value=core.get_kosten(state, uid, typ, merkmal),
                             step=1000.0, format="%.0f",
                             help="Leer lassen, wenn nicht bezifferbar.")
    dauer = st.number_input("Dauer in Wochen (optional)", min_value=0.0,
                            value=core.get_dauer(state, uid, typ, merkmal),
                            step=1.0, format="%.0f",
                            help="Leer lassen, wenn nicht abschaetzbar.")
    sp1, sp2 = st.columns(2)
    if sp1.button("Speichern", type="primary", width="stretch",
                  disabled=(stufe is None)):
        core.set_soll(state, uid, typ, merkmal, ziel_opt)
        core.set_aufwand(state, uid, typ, merkmal, stufe)
        core.set_kosten(state, uid, typ, merkmal, kosten)
        core.set_dauer(state, uid, typ, merkmal, dauer)
        st.rerun()
    if sp2.button("Abbrechen", width="stretch"):
        st.rerun()
    if stufe is None:
        st.caption("Bitte eine Aufwandsstufe wählen, um zu speichern.")


def detail_merkmal_block(state, features, types, uid, merkmal, typ, status):
    """Merkmals-Tabelle in der Detaillierung, bezogen auf EINEN Typ (Tab).
    Knopfregel nach Wissenssituation:
      - Fall 1 (Ist passt): kein Knopf, das Ziel ist automatisch das Ist.
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

    # Kandidaten (Auspraegungen mit Knopf): in Fall 1 keine (Ist ist erfuellt),
    # in allen uebrigen Faellen alle Auspraegungen zur freien Wahl (siehe
    # core.soll_kandidaten).
    kandidaten = core.soll_kandidaten(status, optionen)

    # (Die Soll-Vorbelegung fuer Faelle 1/2/3 erfolgt zentral in
    #  seite_detaillierung via core.soll_vorbelegen, damit auch nicht geoeffnete
    #  Tabs vollstaendig sind.)

    st.markdown(f"**{merkmal}**")

    breiten = [1, 5, 3, 2, 3]
    kopf = st.columns(breiten)
    kopf[0].caption("Profil")
    kopf[1].caption("Ausprägung")
    kopf[2].caption("Ihre Angabe")
    kopf[3].caption("Ziel")
    kopf[4].caption("Aufwand & Kosten")

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

        # Spalte 4: Ziel-Auspraegung (Knopf oeffnet das Aufwand-Pop-up).
        _detail_soll_spalte(state, uid, typ, merkmal, opt, status, soll_wert,
                            ist_wert, kandidaten, profil, z[3])

        # Spalte 5: Aufwand & Kosten der gewaehlten, vom Ist abweichenden Ziel-
        # Auspraegung. Bei Ist=Ziel oder 'nichts anstreben' bleibt sie leer.
        if (opt == soll_wert and opt != ist_wert
                and soll_wert not in (core.OFFEN, core.NICHTS_ANSTREBEN)):
            aufw = core.get_aufwand(state, uid, typ, merkmal)
            kost = core.get_kosten(state, uid, typ, merkmal)
            teile = []
            teile.append(core.AUFWAND_LABEL[aufw] if aufw in core.AUFWAND_LABEL
                         else "⚠️ offen")
            if kost is not None:
                teile.append(f"{kost:,.0f} €".replace(",", "."))
            z[4].markdown("  ·  ".join(teile))
        else:
            z[4].markdown("&nbsp;", unsafe_allow_html=True)

    # In Fall 4 & 5: zusaetzlich der Knopf 'nichts anstreben'.
    soll_final = core.get_soll(state, uid, typ, merkmal)
    if ist_unbekannt_fall:
        _nichts_anstreben_knopf(state, uid, typ, merkmal, soll_final)


def _detail_soll_spalte(state, uid, typ, merkmal, opt, status, soll_wert,
                        ist_wert, kandidaten, profil, spalte):
    """Rechte Spalte (angestrebt) einer Auspraegungszeile. Fall 1: kein Knopf,
    das Ist ist angestrebt. Sonst: Knopf nur fuer Kandidaten. Ausgeschlossene
    Auspraegungen koennen (in Fall 4/5) gewaehlt werden, ohne den Ausschluss in
    Schritt 1 zu aendern."""
    if status == core.STATUS_IST:
        if opt == ist_wert:
            spalte.markdown("● **Ziel**")
        else:
            spalte.markdown("&nbsp;", unsafe_allow_html=True)
        return

    if opt not in kandidaten:
        spalte.markdown("&nbsp;", unsafe_allow_html=True)
        return

    _soll_knopf(state, uid, typ, merkmal, opt, soll_wert, ist_wert, profil, spalte)


def _soll_knopf(state, uid, typ, merkmal, opt, soll_wert, ist_wert, profil, spalte):
    """Wahl-Knopf fuer die Ziel-Auspraegung, fuer EINEN Typ. Kein Toggle:
    die aktive Auspraegung ist ein Marker, inaktive sind waehlbare Knoepfe.
    Weicht die gewaehlte Auspraegung vom Ist ab, oeffnet der Klick das Aufwand-
    Pop-up (Pflicht-Stufe, optionale Kosten). Ist die gewaehlte Auspraegung das
    Ist selbst, wird ohne Pop-up uebernommen, da keine Veraenderung anfaellt."""
    if opt == soll_wert:
        spalte.markdown("● **Ziel**")
    else:
        if spalte.button("○ wählen", key=f"soll_{uid}_{typ}_{merkmal}_{opt}",
                         type="secondary", width="stretch"):
            if opt == ist_wert:
                core.set_soll(state, uid, typ, merkmal, opt)
                core.set_aufwand(state, uid, typ, merkmal, None)
                core.set_kosten(state, uid, typ, merkmal, None)
                st.rerun()
            else:
                _aufwand_dialog(state, uid, typ, merkmal, opt, opt in profil)


def _nichts_anstreben_knopf(state, uid, typ, merkmal, soll_wert):
    """Knopf 'nichts anstreben' fuer die Faelle 4 und 5 (kein Ist). Eine bewusste
    Entscheidung, fuer dieses Merkmal zu diesem Typ nichts anzustreben; sie zaehlt
    im Ziel-Score als Nichttreffer und erfuellt die Pflicht zur Angabe. Toggle:
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
    """Inhalt eines Typ-Tabs in der Detaillierung: Vergleichsfall-Kennzahlen und die fuenf
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
    k1.metric("✅ Passt aktuell", anzahl[core.STATUS_IST])
    k2.metric("🟠 Passt potenziell", anzahl[core.STATUS_POTENTIAL])
    k3.metric("🔶 Passt nicht", anzahl[core.STATUS_IST_UNPASSEND])
    k4.metric("⬜ Offen", anzahl[core.STATUS_OFFEN])
    k5.metric("⛔ Ausgeschlossen", anzahl[core.STATUS_BLOCKIERT])

    # --- Fuenf aufklappbare Abschnitte, je Status einer ---
    abschnitte = [
        (core.STATUS_IST,       "✅ Passt aktuell",
         "Die aktuelle Ausprägung liegt im Profil und wird angestrebt. Hier "
         "besteht kein Handlungsbedarf."),
        (core.STATUS_POTENTIAL, "🟠 Passt potenziell",
         "Eine als Potenzial erfasste Ausprägung liegt im Profil. Wählen Sie sie "
         "als Ziel und schätzen Sie den Aufwand."),
        (core.STATUS_IST_UNPASSEND, "🔶 Passt nicht",
         "Weder die aktuelle noch eine Potenzial-Ausprägung liegt im Profil. "
         "Behalten Sie die aktuelle Ausprägung oder streben Sie eine "
         "Profilausprägung an."),
        (core.STATUS_OFFEN,     "⬜ Offen",
         "Für dieses Merkmal liegt keine Angabe vor. Wählen Sie eine "
         "Ziel-Ausprägung oder „nichts anstreben“."),
        (core.STATUS_BLOCKIERT, "⛔ Ausgeschlossen",
         "Alle Profilausprägungen wurden ausgeschlossen. Prüfen Sie noch einmal, "
         "ob das wirklich zutrifft. Ist eine Ausprägung doch erreichbar, können "
         "Sie sie hier wählen, andernfalls wählen Sie „nichts anstreben“."),
    ]
    for status, titel, hinweis in abschnitte:
        merkmale = gruppen[status]
        with st.expander(f"{titel}  ({len(merkmale)})",
                         expanded=(status in (core.STATUS_POTENTIAL,
                                              core.STATUS_IST_UNPASSEND,
                                              core.STATUS_OFFEN)
                                   and len(merkmale) > 0)):
            if not merkmale:
                st.caption("— keine Merkmale in diesem Vergleichsfall —")
            else:
                st.caption(hinweis)
                for m in merkmale:
                    with st.container(border=True):
                        detail_merkmal_block(state, features, types, uid, m,
                                             typ, status)


HILFE_ZIELTYP = """
**Schritt 3 von 4 · Zieltypbestimmung (Teil 1 von 2)**

Für jeden Typ der engeren Auswahl legen Sie hier fest, welche Ausprägung Sie je
Merkmal künftig anstreben, und schätzen den dafür nötigen Aufwand. So entsteht je
Typ eine vollständige Zielkonfiguration.

*Zielkonfiguration: die Gesamtheit aller Ziel-Ausprägungen einer
Betrachtungseinheit, ausgerichtet am Vorbild eines Typs. Sie beschreibt, wie die
Auftragsabwicklung künftig aussehen soll.*

**1. Wählen Sie den Tab einer Betrachtungseinheit und darin den Tab eines Typs.**

>>**1.1 Legen Sie je Merkmal die Ziel-Ausprägung fest.** Die Merkmale sind nach
denselben fünf Fällen gruppiert wie im Abgleich des vorherigen Schritts. Je Merkmal
gibt es eine Tabelle. Sie zeigt links das Profil des Typs, an dem Sie sich
orientieren, und rechts wählen Sie in der Spalte Ziel die angestrebte Ausprägung.

>>*Ziel-Ausprägung: die Ausprägung, die Sie für dieses Merkmal künftig anstreben.
Je Merkmal und Typ ist genau eine möglich. Merkmale, die bereits passen, brauchen
keine Eingabe. Bei unbekannten Merkmalen können Sie über den Knopf* nichts
anstreben *auch bewusst auf eine Ziel-Ausprägung verzichten, wenn das Merkmal für
Ihren Anwendungsfall nicht sinnvoll festzulegen ist.*

>>**1.2 Weicht die Ziel-Ausprägung von der aktuellen Ausprägung ab,
schätzen Sie im Fenster den Aufwand.**

>>*Aufwandsstufe (Pflicht): wie tief der Eingriff reicht.
Gering bedeutet umsetzbar innerhalb bestehender Prozesse und Systeme, mittel eine
Anpassung bestehender Prozesse oder Systeme, hoch den Aufbau neuer Prozesse,
Systeme oder Ressourcen.*

>>*Kosten und Dauer (freiwillig): beziffern den Aufwand
genauer. Lassen Sie die Felder leer, wenn keine belastbare Schätzung möglich ist.
Erfasste Kosten erlauben später eine budgetbezogene Auswertung.*

>>*Bsp: Der Aufbau einer eigenen Aufbereitungsstätte ist ein
hoher Aufwand, die Umstellung eines Planungsparameters ein geringer.*

**2. Wiederholen Sie das für jeden Typ der engeren Auswahl und für jede
Betrachtungseinheit.**

Sobald für jeden gewählten Typ alle Merkmale entschieden sind, geht es links über
den Knopf zum Typvergleich weiter, wo Sie die Typen gegenüberstellen und den
Zieltyp festlegen.
"""


def seite_detaillierung(state, features, types, weights):
    st.title("Zieltypbestimmung")

    # Das Soll wird in der Seitenleiste (alle_detaillierungen_vollstaendig) fuer
    # die Faelle 1/2/3 vorbelegt; die Seitenleiste laeuft vor dieser Seite.
    if not state["units"]:
        st.caption("Keine Betrachtungseinheiten vorhanden.")
        return

    # Nur Einheiten mit mindestens einem Typ in der engeren Auswahl koennen
    # detailliert werden.
    ohne_auswahl = [uid for uid in state["units"]
                    if not core.get_engere_auswahl(state, uid)]
    if ohne_auswahl:
        st.warning("Für folgende Betrachtungseinheiten ist noch kein Typ in der "
                   f"engeren Auswahl: {', '.join(ohne_auswahl)}. Wählen Sie sie im "
                   "vorherigen Schritt aus.")

    st.divider()
    _tabs = st.tabs([core.get_name(state, uid) for uid in state["units"]])
    for _tab, uid in zip(_tabs, state["units"]):
        with _tab:
            auswahl = core.get_engere_auswahl(state, uid)
            if not auswahl:
                st.caption("Keine Typen in der engeren Auswahl, übersprungen.")
                st.divider()
                continue

            final = core.get_zieltyp(state, uid)
            # Ein Tab je Typ der engeren Auswahl; der finale Zieltyp ist markiert.
            tab_titel = [(f"⭐ {t}" if t == final else t) for t in auswahl]
            for tab, typ in zip(st.tabs(tab_titel), auswahl):
                with tab:
                    _detail_tab_inhalt(state, features, types, uid, typ)

            st.divider()

_STUFEN_TEXT = {core.AUFWAND_KEIN: "ohne", core.AUFWAND_GERING: "gering",
                core.AUFWAND_MITTEL: "mittel", core.AUFWAND_HOCH: "hoch"}


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
        felder = core.handlungsfelder(state, uid, typ, types[typ],
                                      features, weights)
        neu_je_stufe = {k: [] for k in core.AUFWAND_STUFEN_K}
        for f in felder:
            st_k = (f["aufwand"] if f["aufwand"] in core.AUFWAND_STUFEN
                    else core.AUFWAND_HOCH)
            neu_je_stufe[st_k].append(f["merkmal"])
        for k in core.AUFWAND_STUFEN_K:
            zeilen.append({"stufe": k, "aehnlichkeit": score[k], "typ": typ,
                           "stufe_text": _STUFEN_TEXT[k],
                           "neu": ", ".join(neu_je_stufe[k]) or "—"})
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
                                      labelExpr="datum.value == 0 ? 'ohne' : "
                                                "datum.value == 1 ? 'gering' : "
                                                "datum.value == 2 ? 'mittel' : 'hoch'",
                                      title="Aufwand")),
                y=alt.Y("aehnlichkeit:Q",
                        scale=alt.Scale(domain=[0, 1]),
                        axis=alt.Axis(format="%",
                                      title="erreichbare Ähnlichkeit")),
                color=alt.Color("typ:N", title="Typ"),
                tooltip=[alt.Tooltip("typ:N", title="Typ"),
                         alt.Tooltip("stufe_text:N", title="Aufwandsstufe"),
                         alt.Tooltip("neu:N", title="ab dieser Stufe neu"),
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

    st.caption("Stufe 0 ist die ohne jede Änderung erreichte Ähnlichkeit. "
               "Jede weitere Stufe schliesst die vorhergehenden ein, die Verläufe "
               "können daher nicht fallen.")


def _budgetdiagramm(state, features, types, weights, uid, auswahl):
    """Budgetkurve je Betrachtungseinheit: erreichbare Ziel-Aehnlichkeit ueber dem
    kumulierten Budget, ein Treppen-Verlauf je Typ der engeren Auswahl. Grundlage
    sind die Handlungsfelder mit erfassten Kosten. Handlungsfelder ohne Kosten
    erscheinen nicht in der Kurve, sondern werden darunter je Typ als noch nicht
    bezifferbar ausgewiesen; sie deckeln die planbare Kurve."""
    import pandas as pd

    zeilen, infos, hat_kurve = [], [], False
    for typ in auswahl:
        bk = core.budgetkurve(state, uid, typ, types[typ], features, weights)
        for budget, aehn, merkmale in bk["punkte"]:
            zeilen.append({"budget": budget, "aehnlichkeit": aehn, "typ": typ,
                           "merkmale": " + ".join(merkmale) if merkmale else "—"})
        if len(bk["punkte"]) > 1:
            hat_kurve = True
        infos.append((typ, bk))

    if not hat_kurve:
        st.caption("Noch keine Kosten erfasst. Tragen Sie in der "
                   "Zieltypbestimmung Kosten je Änderung ein, damit die "
                   "Budgetkurve entsteht.")
        return

    df = pd.DataFrame(zeilen)
    try:
        import altair as alt
        grund = alt.Chart(df).encode(
            x=alt.X("budget:Q",
                    scale=alt.Scale(domainMin=0, nice=False),
                    axis=alt.Axis(title="Budget (Euro)", format="~s")),
            y=alt.Y("aehnlichkeit:Q",
                    scale=alt.Scale(domain=[0, 1]),
                    axis=alt.Axis(format="%", title="erreichbare Ähnlichkeit")),
            color=alt.Color("typ:N", title="Typ"),
        )
        diagramm = grund.mark_line(interpolate="step-after", point=True,
                                   strokeWidth=2).encode(
            tooltip=[alt.Tooltip("typ:N", title="Typ"),
                     alt.Tooltip("budget:Q", title="Budget", format=",.0f"),
                     alt.Tooltip("aehnlichkeit:Q", title="Ähnlichkeit",
                                 format=".0%"),
                     alt.Tooltip("merkmale:N", title="enthaltene Merkmale")]
            ).properties(height=340)
        try:
            st.altair_chart(diagramm, width="stretch")
        except TypeError:
            st.altair_chart(diagramm, use_container_width=True)
    except Exception:
        st.line_chart(df.pivot(index="budget", columns="typ",
                               values="aehnlichkeit"))

    for typ, bk in infos:
        if bk["unbeziffert"]:
            namen = ", ".join(f["merkmal"] for f in bk["unbeziffert"])
            st.caption(
                f"{typ}: {len(bk['unbeziffert'])} Merkmal(e) ohne Kostenschätzung "
                f"({namen}). Die planbare Kurve deckelt bei "
                f"{bk['deckel'] * 100:.0f} %, mit diesen Änderungen wären bis zu "
                f"{bk['max_gesamt'] * 100:.0f} % erreichbar.")

    if any(not bk.get("exakt", True) for _, bk in infos):
        st.caption("Bei mindestens einem Typ liegen sehr viele bezifferte "
                   "Handlungsfelder vor, dort wird die Kurve näherungsweise "
                   "gerechnet.")
    st.caption("Die Kurve zeigt zu jedem Budget die höchste erreichbare "
               "Ähnlichkeit über alle Kombinationen der bezifferten Änderungen. "
               "Beim Überfahren eines Punktes erscheinen die darin enthaltenen "
               "Merkmale.")


def _dauer_diagramm(state, features, types, weights, uid, auswahl):
    """Dauer je Aenderung als Balken, die alle bei null beginnen. Das entspricht
    der parallelen Sicht: laufen die Aenderungen gleichzeitig, bestimmt die
    laengste Dauer die Gesamtdauer. Ein Balken je Merkmal und Typ."""
    import pandas as pd
    zeilen = []
    for typ in auswahl:
        for f in core.handlungsfelder(state, uid, typ, types[typ],
                                      features, weights):
            d = core.get_dauer(state, uid, typ, f["merkmal"])
            if d is not None:
                zeilen.append({"merkmal": f["merkmal"], "typ": typ, "dauer": d,
                               "label": f"{f['merkmal']} · {typ}"})
    if not zeilen:
        st.caption("Sobald Dauern erfasst sind, erscheint hier je Änderung ein "
                   "Balken, der bei null beginnt.")
        return
    df = pd.DataFrame(zeilen)
    try:
        import altair as alt
        chart = (alt.Chart(df).mark_bar().encode(
            x=alt.X("dauer:Q", scale=alt.Scale(domainMin=0, nice=False),
                    axis=alt.Axis(title="Dauer (Wochen)")),
            y=alt.Y("label:N", sort="-x", title=None),
            color=alt.Color("typ:N", title="Typ"),
            tooltip=[alt.Tooltip("merkmal:N", title="Merkmal"),
                     alt.Tooltip("typ:N", title="Typ"),
                     alt.Tooltip("dauer:Q", title="Dauer (Wochen)")])
            .properties(height=60 + 44 * len(df)))
        try:
            st.altair_chart(chart, width="stretch")
        except TypeError:
            st.altair_chart(chart, use_container_width=True)
    except Exception:
        st.bar_chart(df.set_index("label")["dauer"])
    st.caption("Jeder Balken beginnt bei null. Laufen die Änderungen parallel, "
               "bestimmt die längste Dauer die Gesamtdauer.")


def _typ_merkmal_tabelle(state, uid, typ, profil, features, weights):
    """Uebersicht der Merkmale mit Aenderungsbedarf fuer EINEN Typ: genau die
    Deltas, die als Stufen (Aufwandsdiagramm) bzw. Spruenge (Budgetkurve) in die
    Diagramme eingehen. Ohne Synergiepotenzial, das gehoert in die
    Konsolidierung."""
    import pandas as pd
    felder = core.handlungsfelder(state, uid, typ, profil, features, weights)
    if not felder:
        st.caption("Keine Änderung nötig, der Ist-Zustand entspricht bereits dem "
                   "Profil dieses Typs.")
        return
    zeilen = []
    for f in felder:
        kost = core.get_kosten(state, uid, typ, f["merkmal"])
        dau = core.get_dauer(state, uid, typ, f["merkmal"])
        zeilen.append({
            "Merkmal": f["merkmal"],
            "Ist": _fmt_auspr(f["ist"]),
            "Ziel": f["soll"],
            "Ähnlichkeitsgewinn": f"+{f['gewinn'] * 100:.1f} %-Pkt.",
            "Aufwand": core.AUFWAND_LABEL.get(f["aufwand"], "—"),
            "Kosten": (f"{kost:,.0f} €".replace(",", ".")
                       if kost is not None else "—"),
            "Dauer": (f"{dau:.0f} Wo." if dau is not None else "—"),
        })
    st.dataframe(pd.DataFrame(zeilen), hide_index=True, width="stretch")


def _zieltyp_festlegung(state, uid, auswahl):
    """Festlegung des finalen Zieltyps als Einzelauswahl-Dropdown, analog zum
    Auswahlfeld der Typenvorauswahl. Die erste Option hebt die Festlegung auf."""
    st.subheader("Zieltyp festlegen")
    final = core.get_zieltyp(state, uid)
    KEINE = "— noch nicht festgelegt —"
    optionen = [KEINE] + list(auswahl)
    index = optionen.index(final) if final in auswahl else 0
    wahl = st.selectbox(
        f"Zieltyp für Betrachtungseinheit {core.get_name(state, uid)}",
        options=optionen, index=index, key=f"zieltyp_wahl_{uid}",
        help="Der Zieltyp ist die verbindliche Grundlage der Transformationsplanung.")
    neuer = None if wahl == KEINE else wahl
    if neuer != final:
        core.set_zieltyp(state, uid, neuer)
        st.rerun()


HILFE_TYPVERGLEICH = """
**Schritt 3 von 4 · Typvergleich (Teil 2 von 2)**

Hier stellen Sie die ausgearbeiteten Zielkonfigurationen gegenüber und legen je
Betrachtungseinheit den Zieltyp fest. Das Werkzeug zeigt dafür, welche Ähnlichkeit
mit welchem Aufwand, welchem Budget und welcher Dauer erreichbar ist.

**Aufbau der Seite**

**1. Tab je Betrachtungseinheit.**

**2. Merkmale mit Änderungsbedarf.** Je Typ eine Tabelle mit den Änderungen, die
nötig wären, um sich diesem Typ anzunähern. Aufgeführt sind alle Merkmale, deren
Ziel-Ausprägung von der aktuellen abweicht, mit dem geschätzten Aufwand aus
Aufwandsstufe und Kosten sowie dem Nutzen in Form des Ähnlichkeitsgewinns.

*Ähnlichkeitsgewinn: um wie viele Prozentpunkte die Ähnlichkeit zum Typ steigt,
wenn diese Änderung umgesetzt wird.*

**3. Diagramm: Erreichbare Ähnlichkeit nach Aufwand.** Je Typ ein Verlauf über die
vier Aufwandsstufen. Er zeigt die Ähnlichkeit zum Typ, wenn alle Änderungen bis zu
dieser Stufe umgesetzt werden. Ohne bedeutet keine Änderung, gering nur die
geringen, mittel zusätzlich die mittleren und hoch alle.

Welcher Typ am besten passt, kann sich mit der Aufwandsstufe ändern. So kann ein
Typ ohne jede Änderung die höchste Ähnlichkeit aufweisen, während ein anderer ihn
bei mittlerem Aufwand übertrifft.

**4. Diagramm: Erreichbare Ähnlichkeit nach Budget.** Je Typ eine Kurve, die zu jedem Budget
die höchste erreichbare Ähnlichkeit zeigt. Zudem, welche Kombination von Änderungen
dafür umzusetzen ist. Änderungen ohne Kostenangabe fließen nicht ein und begrenzen
die Kurve nach oben.

**5. Dauer der Änderungen.** Ein Balken je Änderung, alle beginnen bei null. Bei
paralleler Umsetzung bestimmt die längste Dauer die Gesamtdauer.

**6. Zieltyp festlegen.** Wählen Sie je Betrachtungseinheit einen Zieltyp.

*Zieltyp: der Typ, an dem Sie die weitere Ausgestaltung der Betrachtungseinheit
orientieren. Die gewählte Zielkonfiguration und die dadurch nötigen Änderungen sind
die Grundlage für den nächsten Schritt.*

Sobald jede Betrachtungseinheit einen Zieltyp hat, geht es links über den Knopf zur
Konsolidierung weiter.
"""


def seite_zusammenfassung(state, features, types, weights):
    st.title("Typvergleich")

    if not state["units"]:
        st.caption("Keine Betrachtungseinheiten vorhanden.")
        return

    st.divider()
    _tabs = st.tabs([core.get_name(state, uid) for uid in state["units"]])
    for _tab, uid in zip(_tabs, state["units"]):
        with _tab:
            auswahl = core.get_engere_auswahl(state, uid)
            if not auswahl:
                st.caption("Keine Typen in der engeren Auswahl.")
                st.divider()
                continue

            final = core.get_zieltyp(state, uid)

            # 1. Merkmale mit Aenderungsbedarf je Typ (die Deltas, die als
            #    Stufen bzw. Spruenge in die beiden Diagramme eingehen).
            st.subheader("Merkmale mit Änderungsbedarf")
            for typ in auswahl:
                marke = "⭐ " if typ == final else ""
                st.markdown(f"### {marke}{typ}")
                _typ_merkmal_tabelle(state, uid, typ, types[typ],
                                     features, weights)

            # 2. Zwei Sichten auf die erreichbare Aehnlichkeit: nach
            #    Aufwandsstufe (Rueckgrat) und nach Budget (kostenbewusst).
            st.markdown("**Erreichbare Ähnlichkeit nach Aufwand**")
            _stufendiagramm(state, features, types, weights, uid, auswahl)
            st.markdown("**Erreichbare Ähnlichkeit nach Budget**")
            _budgetdiagramm(state, features, types, weights, uid, auswahl)

            # 3. Budget-Loeser: bester Typ zu einem konkret eingegebenen Budget
            #    (exaktes Optimum statt gieriger Kurve).
            # 3b. Dauer der Aenderungen als Balken (alle bei null beginnend).
            st.markdown("**Dauer der Änderungen**")
            _dauer_diagramm(state, features, types, weights, uid, auswahl)

            # 4. Zieltyp festlegen: Einzelauswahl-Dropdown ganz unten, nachdem
            #    Tabellen und Diagramme die Grundlage geliefert haben.
            _zieltyp_festlegung(state, uid, auswahl)
            st.divider()




def _fmt_auspr(wert):
    """Anzeige einer Auspraegung; ein unbekanntes Ist erscheint als Gedankenstrich."""
    return "—" if wert is core.OFFEN else str(wert)


def _handlungsfeld_tabelle(state, uid, felder):
    """Automatisch erzeugte Uebersicht der Handlungsfelder je Betrachtungseinheit.
    Die Spalte Synergiepotenzial nennt andere Einheiten mit gleicher Ziel-Ausprägung
    und ist reine Anzeige, die Zusammenfuehrung erfolgt im naechsten Schritt."""
    import pandas as pd
    zeilen = []
    hat_synergie = False
    hat_vorbild = False
    for f in felder:
        treffer = core.synergie_potenzial(state, uid, f["merkmal"], f["soll"])
        teile = []
        for uid2, besitzt in treffer:
            teile.append(f"{uid2}*" if besitzt else uid2)
            hat_vorbild = hat_vorbild or besitzt
        if treffer:
            hat_synergie = True
        zeilen.append({
            "Merkmal": f["merkmal"],
            "Ist": _fmt_auspr(f["ist"]),
            "Ziel": f["soll"],
            "Aufwand": core.AUFWAND_LABEL.get(f["aufwand"], "—"),
            "Ähnlichkeitsgewinn": f"+{f['gewinn'] * 100:.1f} %-Pkt.",
            "Synergiepotenzial": ", ".join(teile) if teile else "—",
        })
    st.dataframe(pd.DataFrame(zeilen), hide_index=True, width="stretch")
    if hat_synergie:
        txt = ("Synergiepotenzial nennt weitere Betrachtungseinheiten, die im selben "
               "Merkmal dieselbe Ziel-Ausprägung verfolgen.")
        if hat_vorbild:
            txt += (" Ein Sternchen kennzeichnet eine Einheit, welche die Ausprägung "
                    "bereits besitzt und damit als Vorbild dienen kann.")
        st.caption(txt)


def _portfolio_matrix(punkte):
    """Aufwand-Wirkung-Portfolio: Aufwand auf der Abszisse, Ähnlichkeitsgewinn auf
    der Ordinate. Handlungsfelder mit hohem Gewinn und geringem Aufwand liegen links
    oben und sind die naheliegenden ersten Schritte. Erwartet je Punkt ein dict mit
    label, aufwand und gewinn (Anteil)."""
    import pandas as pd
    if not punkte:
        return
    # Punkte mit gleichem Aufwand UND gleichem Gewinn liegen deckungsgleich. Sie
    # werden zu einem Punkt zusammengefasst und gemeinsam beschriftet, damit sich
    # die Beschriftungen nicht ueberdecken.
    gruppen = {}
    for p in punkte:
        schluessel = (p["aufwand"] or core.AUFWAND_HOCH, round(p["gewinn"], 6))
        gruppen.setdefault(schluessel, []).append(p["label"])
    zeilen = [{"aufwand": a, "gewinn": g * 100,
               "label": ", ".join(namen), "anzahl": len(namen)}
              for (a, g), namen in gruppen.items()]
    df = pd.DataFrame(zeilen)
    try:
        import altair as alt
        basis = alt.Chart(df).encode(
            x=alt.X("aufwand:Q", scale=alt.Scale(domain=[0.5, 3.5], nice=False),
                    axis=alt.Axis(values=[1, 2, 3], title="Aufwand",
                                  labelExpr="datum.value == 1 ? 'gering' : "
                                            "datum.value == 2 ? 'mittel' : 'hoch'")),
            y=alt.Y("gewinn:Q", scale=alt.Scale(zero=True),
                    axis=alt.Axis(title="Ähnlichkeitsgewinn (Prozentpunkte)")),
        )
        kreise = basis.mark_circle(opacity=0.7).encode(
            size=alt.Size("anzahl:Q", legend=None,
                          scale=alt.Scale(domain=[1, 6], range=[160, 420])),
            tooltip=[alt.Tooltip("label:N", title="Handlungsfeld"),
                     alt.Tooltip("aufwand:Q", title="Aufwand"),
                     alt.Tooltip("gewinn:Q", title="Ähnlichkeitsgewinn (%-Pkt.)",
                                 format=".1f")],
        )
        beschriftung = basis.mark_text(
            align="left", dx=12, fontSize=11, baseline="middle").encode(text="label:N")
        diagramm = (kreise + beschriftung).properties(height=340)
        try:
            st.altair_chart(diagramm, width="stretch")
        except TypeError:
            st.altair_chart(diagramm, use_container_width=True)
    except Exception:
        st.dataframe(df, hide_index=True, width="stretch")
    st.caption(
        "Die Reihenfolge richtet sich zuerst nach dem Ähnlichkeitsgewinn und bei "
        "gleichem Gewinn nach dem Aufwand. Zuerst kommen Handlungsfelder mit hohem "
        "Gewinn und geringem Aufwand, danach solche mit hohem Gewinn und hohem "
        "Aufwand, anschliessend solche mit geringem Gewinn und geringem Aufwand und "
        "zuletzt solche mit geringem Gewinn und hohem Aufwand. Punkte mit gleichem "
        "Aufwand und Gewinn sind zusammengefasst.")


def _paretokurve(punkte):
    """Kumulative Paretokurve: die Handlungsfelder nach absteigendem Ähnlichkeits-
    gewinn geordnet, aufgetragen ist der aufsummierte Gewinn. Die Kurve steigt
    zuerst steil und flacht ab, sodass sichtbar wird, dass wenige Handlungsfelder den
    Grossteil der Annaeherung tragen."""
    import pandas as pd
    if not punkte:
        return
    sortiert = sorted(punkte, key=lambda p: -p["gewinn"])
    zeilen = []
    kum = 0.0
    for i, p in enumerate(sortiert, 1):
        kum += p["gewinn"] * 100
        zeilen.append({"Rang": i, "kumuliert": round(kum, 2), "label": p["label"]})
    df = pd.DataFrame(zeilen)
    try:
        import altair as alt
        linie = alt.Chart(df).mark_line(point=True).encode(
            x=alt.X("Rang:Q", scale=alt.Scale(domain=[0.5, len(zeilen) + 0.5]),
                    axis=alt.Axis(title="Handlungsfelder nach Gewinn geordnet",
                                  tickMinStep=1, format="d")),
            y=alt.Y("kumuliert:Q", scale=alt.Scale(zero=True),
                    axis=alt.Axis(title="kumulierter Ähnlichkeitsgewinn (%-Pkt.)")),
            tooltip=[alt.Tooltip("label:N", title="Handlungsfeld"),
                     alt.Tooltip("Rang:Q", title="Rang"),
                     alt.Tooltip("kumuliert:Q", title="kumuliert (%-Pkt.)",
                                 format=".1f")])
        diagramm = linie.properties(height=300)
        try:
            st.altair_chart(diagramm, width="stretch")
        except TypeError:
            st.altair_chart(diagramm, use_container_width=True)
    except Exception:
        st.dataframe(df, hide_index=True, width="stretch")
    st.caption("Wenige hoch gewichtete Handlungsfelder tragen den Grossteil der "
               "Annaeherung an die Zieltypen. Die flach auslaufenden Handlungsfelder "
               "sind nachrangig.")


def _morphologie_html(state, uid, typ_name, typ_profil, features):
    """Morphologischer Kasten fuer den Export: alle Merkmale mit ihren
    Auspraegungen, das Profil des Zieltyps grau hinterlegt.

    Die Merkmale sind nach der Erfassungsregel in zwei Bloecke geteilt. Bei
    bekanntem Ist-Zustand werden die Ist- und die Ziel-Auspraegung gekennzeichnet,
    bei Uebereinstimmung beider als Ist = Ziel. Bei unbekanntem Ist-Zustand wird
    allein die Ziel-Auspraegung ausgewiesen."""
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
                    marke = "Ist = Ziel"
            elif j == i_ist:
                marke = "Ist"
            elif j == i_soll:
                marke = "Ziel"
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
                 "Bei bekannter Ausprägung sind die Ist- und die Ziel-Ausprägung "
                 "gekennzeichnet, bei Übereinstimmung beider als Ist = Ziel.</p>")
    return "\n".join(teile)


_BE_FARBEN = ["#2a78d6", "#d64545", "#2f9e44", "#e8590c", "#7048e8",
              "#0c8599", "#c2255c", "#5c940d", "#1864ab", "#a61e4d"]


def _be_farben(einheiten):
    """Ordnet jeder Betrachtungseinheit eine feste Farbe aus der Palette zu."""
    return {uid: _BE_FARBEN[i % len(_BE_FARBEN)] for i, uid in enumerate(einheiten)}


def _be_legende_html(state, einheiten, farben):
    """Farblegende der Einheiten mit Erklaerung der Ist-/Ziel-Kennzeichnung."""
    teile = ["<p class='leg'><b>Betrachtungseinheiten:</b> "]
    for uid in einheiten:
        teile.append(f"<span class='bg z' style='background:{farben[uid]}'>"
                     f"{core.get_name(state, uid)}</span> ")
    teile.append("</p>")
    # Je Legendenelement eine eigene Zeile, mit einer Beispielmarke in der
    # Darstellung, die auch im Kasten verwendet wird.
    bsp = einheiten[0] if einheiten else None
    farbe = farben[bsp] if bsp else "#333333"
    teile.append("<p class='leg'><span class='sub'>"
                 f"<span class='bg z' style='background:{farbe}'>Name</span> "
                 "Ziel-Ausprägung<br>"
                 f"<span class='bg i' style='color:{farbe};border-color:{farbe}'>"
                 "Name</span> Ist-Ausprägung<br>"
                 f"<span class='bg z' style='background:{farbe}'>= Name</span> "
                 "Ist-Ausprägung entspricht bereits der Ziel-Ausprägung"
                 "</span></p>")
    return "".join(teile)


def _gemeinsamer_kasten_html(state, einheiten, features, types, farben):
    """Ein morphologischer Kasten fuer alle Einheiten. Je Merkmal und Auspraegung
    werden die Einheiten als farbige Marken eingetragen: Ziel gefuellt, Ist
    umrandet, Ist gleich Ziel gefuellt mit vorangestelltem Gleichheitszeichen."""
    spalten = max((len(features[m]) for m in features), default=0)
    teile = ["<table class='morph'>"]
    for m in features:
        optionen = list(features[m])
        teile.append(f"<tr><th class='mn'>{m}</th>")
        for opt in optionen:
            marken = []
            for uid in einheiten:
                typ = core.get_zieltyp(state, uid)
                ist = core.get_choice(state, uid, m)
                soll = core.get_soll(state, uid, typ, m) if typ else core.OFFEN
                if soll is core.OFFEN or soll == core.NICHTS_ANSTREBEN:
                    soll = None
                nm = core.get_name(state, uid)
                c = farben[uid]
                if ist == opt and soll == opt:
                    marken.append(f"<span class='bg z' style='background:{c}'>= {nm}</span>")
                elif soll == opt:
                    marken.append(f"<span class='bg z' style='background:{c}'>{nm}</span>")
                elif ist == opt:
                    marken.append(f"<span class='bg i' style='color:{c};"
                                  f"border-color:{c}'>{nm}</span>")
            inner = f"<div class='bgs'>{''.join(marken)}</div>" if marken else ""
            teile.append(f"<td>{opt}{inner}</td>")
        teile.extend("<td class='leer'></td>" for _ in range(spalten - len(optionen)))
        teile.append("</tr>")
    teile.append("</table>")
    return "".join(teile)


def _export_html(state, features, types, weights):
    """Erzeugt das Gesamtergebnis als HTML. Aufbau: je Einheit der Zieltyp und die
    erreichbare Aehnlichkeit je Aufwandsstufe, die Priorisierung der Einheiten, die
    harmonisierte Massnahmentabelle und ein gemeinsamer morphologischer Kasten mit
    allen Einheiten. Der Anwender kann daraus ueber die Druckfunktion seines
    Browsers ein PDF erstellen."""
    einheiten = state["units"]
    farben = _be_farben(einheiten)
    stufen = [(core.AUFWAND_KEIN, "ohne"), (core.AUFWAND_GERING, "gering"),
              (core.AUFWAND_MITTEL, "mittel"), (core.AUFWAND_HOCH, "hoch")]

    def dd(xs):
        return " + ".join(dict.fromkeys(x for x in xs if x)) or "—"

    teile = ["<html><head><meta charset='utf-8'><title>Konfigurationsergebnis</title>",
             "<style>body{font-family:sans-serif;margin:2cm;line-height:1.5}",
             "table{border-collapse:collapse;width:100%;margin:1em 0}",
             "th,td{border:1px solid #999;padding:6px;text-align:left;font-size:10pt}",
             "th{background:#eee}h1{font-size:18pt}h2{font-size:14pt;margin-top:2em}",
             "table.morph td{text-align:center;font-size:9pt;vertical-align:top}",
             "table.morph td.leer{background:#fafafa;border:none}",
             "table.morph th.mn{width:16%;font-weight:600;background:#f4f4f4;text-align:left}",
             "div.bgs{margin-top:3px}",
             "span.bg{display:inline-block;padding:1px 5px;margin:1px;border-radius:3px;"
             "font-size:8pt;font-weight:600;white-space:nowrap}",
             "span.bg.z{color:#fff}span.bg.i{background:#fff;border:1.5px solid}",
             "p.leg{font-size:10pt;color:#333}.sub{font-size:8.5pt;color:#555}",
             "</style></head><body>",
             "<h1>Ergebnis des Konfigurationsvorgehens</h1>"]

    # 1. Zieltyp und erreichbare Aehnlichkeit je Aufwandsstufe.
    teile.append("<h2>Zieltypen und erreichbare Ähnlichkeit</h2>")
    teile.append("<table><tr><th>Betrachtungseinheit</th><th>Zieltyp</th>"
                 + "".join(f"<th>Aufwand {n}</th>" for _, n in stufen) + "</tr>")
    for uid in einheiten:
        typ = core.get_zieltyp(state, uid)
        nm = core.get_name(state, uid)
        if not typ:
            teile.append(f"<tr><td>{nm}</td><td>—</td>"
                         + "".join("<td>—</td>" for _ in stufen) + "</tr>")
            continue
        score = core.soll_score_gestaffelt(state, uid, typ, types[typ], features, weights)
        teile.append(f"<tr><td>{nm}</td><td>{typ}</td>"
                     + "".join(f"<td>{score[k] * 100:.0f} %</td>" for k, _ in stufen)
                     + "</tr>")
    teile.append("</table>")

    # 2. Priorisierung der Einheiten (Rangfolge nach Nutzwert).
    nw = core.nwa_nutzwerte(state, types, features, weights)
    if nw:
        teile.append("<h2>Priorisierung der Betrachtungseinheiten</h2>")
        teile.append("<p>Reihenfolge, in der die Betrachtungseinheiten angegangen "
                     "werden sollten, nach der Nutzwertanalyse.</p>")
        teile.append("<table><tr><th>Rang</th><th>Betrachtungseinheit</th>"
                     "<th>Nutzwert</th></tr>")
        for rang, (uid, wert) in enumerate(nw, 1):
            teile.append(f"<tr><td>{rang}</td><td>{core.get_name(state, uid)}</td>"
                         f"<td>{wert * 100:.1f} %</td></tr>")
        teile.append("</table>")

    # 3. Harmonisierte Massnahmen.
    massnahmen = core.massnahmen_liste(state, types, features, weights)
    if massnahmen:
        teile.append("<h2>Harmonisierte Änderungen</h2>")
        teile.append("<table><tr><th>Nr.</th><th>Merkmal</th>"
                     "<th>Betrachtungseinheit(en)</th>"
                     "<th>Ist</th><th>Ziel</th><th>Ähnlichkeitsgewinn</th>"
                     "<th>Aufwand</th><th>Kosten</th><th>Dauer</th>"
                     "<th>Voraussetzung</th></tr>")
        for i, m in enumerate(massnahmen, 1):
            det = core.get_hf_massnahme(state, m["id"])
            kt = (f"{m['kosten']:,.0f} €".replace(",", ".")
                  if m["kosten"] is not None else "—")
            dt = f"{m['dauer']:.0f} Wo." if m["dauer"] is not None else "—"
            einh = ", ".join(core.get_name(state, u) for u in m["einheiten"])
            teile.append(
                f"<tr><td>{i}</td><td>{dd(m['merkmale'])}</td><td>{einh}</td>"
                f"<td>{dd(m['ist_texte'])}</td><td>{dd(m['ziele'])}</td>"
                f"<td>+{m['gewinn'] * 100:.1f} %-Pkt.</td>"
                f"<td>{core.AUFWAND_LABEL.get(m['aufwand'], '—')}</td>"
                f"<td>{kt}</td><td>{dt}</td>"
                f"<td>{det.get('voraussetzung', '') or '—'}</td></tr>")
        teile.append("</table>")

    # 4. Gemeinsamer morphologischer Kasten mit allen Einheiten.
    teile.append("<h2>Morphologischer Kasten mit allen Betrachtungseinheiten</h2>")
    teile.append(_be_legende_html(state, einheiten, farben))
    teile.append(_gemeinsamer_kasten_html(state, einheiten, features, types, farben))

    teile.append("</body></html>")
    return "".join(teile)


def _ist_anzeige(ist_werte):
    """Ist-Spalte eines unternehmensweiten Handlungsfeldes. Sind alle Ist-Werte
    gleich, wird einer gezeigt, sonst die einzelnen Werte parallel zu den Einheiten."""
    if len(set(ist_werte)) == 1:
        return _fmt_auspr(ist_werte[0])
    return ", ".join(_fmt_auspr(iw) for iw in ist_werte)


def _hf_key_str(k):
    """Stringform eines Handlungsfeld-Schluessels fuer Streamlit-Widget-Keys."""
    return "hf_" + "_".join(str(x) for x in k)


def _uw_zeilen(state, features, types, weights):
    """Unternehmensweite Handlungsfeld-Zeilen, oder None wenn keine Einheit einen
    Zieltyp hat. Die Zeilen tragen auch bei nur einer Einheit, das Zusammenlegen
    (Harmonisierung) setzt jedoch mehrere Einheiten voraus."""
    mit_ziel = [u for u in state["units"] if core.get_zieltyp(state, u)]
    if not mit_ziel:
        return None
    return core.unternehmensweite_uebersicht(state, types, features, weights)


@st.dialog("Gemeinsame Werte der Zusammenlegung")
def _buendel_dialog(state, merkmal, soll, einheiten):
    """Erfasst beim Zusammenlegen gleicher Ziele die gemeinsamen Werte. Der
    Aehnlichkeitsgewinn addiert sich automatisch, hier geht es nur um Aufwand,
    Kosten und Dauer der einmal gemeinsam umgesetzten Aenderung. Vorschlag ist
    jeweils der hoechste Einzelwert der beteiligten Einheiten."""
    st.markdown(f"**{merkmal} → {soll}**")
    st.caption("Beteiligt: " + ", ".join(core.get_name(state, u) for u in einheiten))
    au, ks, ds = core.einzelwerte(state, einheiten, merkmal)
    bw = core.get_buendel_werte(state, merkmal, soll) or {}
    v_auf = bw.get("aufwand", max(au) if au else None)
    v_kos = bw.get("kosten", max(ks) if ks else None)
    v_dau = bw.get("dauer", max(ds) if ds else None)
    st.caption("Vorschlag ist der höchste Einzelwert der beteiligten Einheiten.")
    stufen = list(core.AUFWAND_STUFEN)
    idx = stufen.index(v_auf) if v_auf in stufen else None
    aufwand = st.radio("Gemeinsamer Aufwand", stufen,
                       format_func=lambda x: core.AUFWAND_LABEL[x],
                       index=idx, horizontal=True)
    kosten = st.number_input("Gemeinsame Kosten in Euro (optional)", min_value=0.0,
                             value=v_kos, step=1000.0, format="%.0f")
    dauer = st.number_input("Gemeinsame Dauer in Wochen (optional)", min_value=0.0,
                            value=v_dau, step=1.0, format="%.0f")
    sp1, sp2 = st.columns(2)
    if sp1.button("Speichern", type="primary", width="stretch"):
        core.set_buendel_werte(state, merkmal, soll, aufwand, kosten, dauer)
        st.rerun()
    if sp2.button("Abbrechen", width="stretch"):
        st.rerun()


def _harmonisierung(state, features, types, weights):
    """Vierter Schritt, Teil 1: unternehmensweite Harmonisierung. Die Aenderungen
    aller Einheiten werden als durchnummerierte Massnahmen gefuehrt. Ueber ein
    Auswahlfeld lassen sich beliebige Massnahmen zu einer gemeinsamen zusammenlegen,
    sei es wegen gleicher Ziele oder gleicher Voraussetzungen. Gleiche Ziele werden
    zusaetzlich als Textvorschlag hervorgehoben. Die Voraussetzung je Massnahme ist
    direkt editierbar."""
    import pandas as pd
    zeilen = _uw_zeilen(state, features, types, weights)
    if zeilen is None or not zeilen:
        return
    massnahmen = core.massnahmen_liste(state, types, features, weights)
    st.header("Harmonisierung")

    def _dedup(xs):
        return " + ".join(dict.fromkeys(x for x in xs if x)) or "—"

    daten = []
    for i, m in enumerate(massnahmen, 1):
        det = core.get_hf_massnahme(state, m["id"])
        daten.append({
            "Änderung": str(i),
            "Merkmal": _dedup(m["merkmale"]),
            "Betrachtungseinheit(en)": ", ".join(core.get_name(state, u)
                                                 for u in m["einheiten"]),
            "Ist": _dedup(m["ist_texte"]),
            "Ziel": _dedup(m["ziele"]),
            "Ähnlichkeitsgewinn": f"+{m['gewinn'] * 100:.1f} %-Pkt.",
            "Aufwand": core.AUFWAND_LABEL.get(m["aufwand"], "—"),
            "Kosten": (f"{m['kosten']:,.0f} €".replace(",", ".")
                       if m["kosten"] is not None else "—"),
            "Dauer": (f"{m['dauer']:.0f} Wo." if m["dauer"] is not None else "—"),
            "Voraussetzung": det.get("voraussetzung", ""),
        })
    df = pd.DataFrame(daten)
    fest = ["Änderung", "Merkmal", "Betrachtungseinheit(en)", "Ist", "Ziel",
            "Ähnlichkeitsgewinn", "Aufwand", "Kosten", "Dauer"]
    edit = st.data_editor(
        df, hide_index=True, width="stretch",
        key=f"harm_editor_{len(massnahmen)}",
        column_config={
            **{c: st.column_config.TextColumn(disabled=True) for c in fest},
            "Voraussetzung": st.column_config.TextColumn(
                help="Was vorliegen muss, damit die Änderung umgesetzt werden "
                     "kann: etwa ein ERP-Modul, eine Fläche am Standort oder eine "
                     "Mitarbeiterqualifizierung."),
        })

    def _txt(w):
        return str(w) if pd.notna(w) else ""

    for i, m in enumerate(massnahmen):
        core.set_hf_massnahme(state, m["id"],
                              voraussetzung=_txt(edit.iloc[i]["Voraussetzung"]))

    # Textvorschlag: gleiche Ziele ueber mehrere Einheiten, noch nicht zusammengelegt.
    teile = []
    for (merkmal, soll), einheiten in core.buendel_kandidaten(
            state, types, features, weights).items():
        schluessel = [("e", u, merkmal) for u in einheiten]
        if all(core.synergie_von(state, k) is None for k in schluessel):
            namen = ", ".join(core.get_name(state, u) for u in einheiten)
            teile.append(f"„{merkmal} → {soll}“ ({namen})")
    if teile:
        st.info("Gleiche Ziele bieten sich zum Zusammenlegen an: " + "; ".join(teile)
                + ". Wählen Sie die betreffenden Änderungen unten aus.")

    # Ein Werkzeug: Mehrfachauswahl zum Zusammenlegen beliebiger Massnahmen.
    st.markdown(
        "**Änderungen zusammenlegen** (optional)",
        help="Wählen Sie mehrere Änderungen aus, die sich gemeinsam umsetzen "
             "lassen, etwa weil mehrere Betrachtungseinheiten dasselbe Ziel "
             "verfolgen oder weil verschiedene Änderungen dieselbe Voraussetzung "
             "benötigen. Der Ähnlichkeitsgewinn addiert sich, während Aufwand, "
             "Kosten und Dauer nur einmal anfallen und gemeinsam neu zu schätzen "
             "sind.")
    label = {}
    for i, m in enumerate(massnahmen, 1):
        if not m["ist_synergie"]:
            einh = ", ".join(core.get_name(state, u) for u in m["einheiten"])
            label[f"{i}: {_dedup(m['merkmale'])} → {_dedup(m['ziele'])} ({einh})"] = m["id"]
    auswahl = st.multiselect("Änderungen für eine gemeinsame Änderung wählen",
                             options=list(label), key="harm_select",
                             placeholder="Änderungen wählen")
    if len(auswahl) >= 2:
        if st.button("Zusammenlegen", type="primary", key="harm_merge"):
            _synergie_dialog(state, [label[a] for a in auswahl], massnahmen)
    elif len(auswahl) == 1:
        st.caption("Bitte mindestens zwei Änderungen wählen.")

    synergien = [m for m in massnahmen if m["ist_synergie"]]
    if synergien:
        st.caption("Zusammengelegte Änderungen wieder auflösen:")
        for m in synergien:
            sp1, sp2 = st.columns([5, 1])
            sp1.markdown(f"- **{_dedup(m['merkmale'])}**")
            if sp2.button("Auflösen", key=f"harm_del_{_id_str(m['id'])}"):
                core.remove_synergie_by_id(state, m["id"])
                st.rerun()


def _massnahme_label(state, m):
    """Kurzbezeichnung einer Massnahme aus ihren Merkmalen und Einheiten."""
    einh = ", ".join(core.get_name(state, u) for u in m["einheiten"])
    return f"{' + '.join(m['merkmale'])} ({einh})"


def _id_str(mid):
    """Deterministische Stringform einer Massnahmen-id fuer Streamlit-Widget-Keys."""
    if isinstance(mid, frozenset):
        return "syn__" + "__".join(sorted("-".join(str(x) for x in k) for k in mid))
    return "ein__" + "-".join(str(x) for x in mid)


def _abhaengigkeitsanalyse(state, massnahmen):
    """Paarweise Erhebung zwingender Abhaengigkeiten zwischen den Massnahmen. Je Paar
    legt der Anwender fest, ob eine Massnahme zwingend vor der anderen liegen muss.
    Diese Zwaenge schraenken die spaetere Reihenfolge hart ein."""
    import itertools
    st.subheader("Abhängigkeitsanalyse")
    if len(massnahmen) < 2:
        st.caption("Für eine Abhängigkeitsanalyse sind mindestens zwei Änderungen "
                   "nötig.")
        return
    st.caption("Legen Sie für jedes Paar fest, ob eine Änderung zwingend vor einer "
               "anderen liegen muss, etwa weil sie deren Voraussetzung schafft. Ohne "
               "solche Zwänge bestimmen allein Kosten-Nutzen und Pareto die "
               "Reihenfolge.")
    OPT = ["keine", "a_vor_b", "b_vor_a"]
    for a, b in itertools.combinations(massnahmen, 2):
        la, lb = _massnahme_label(state, a), _massnahme_label(state, b)
        status = core.abhaengigkeit_status(state, a["id"], b["id"])
        wahl = st.radio(
            f"{la}   ·   {lb}", options=OPT, index=OPT.index(status),
            key=f"dep_{_id_str(a['id'])}___{_id_str(b['id'])}",
            format_func=lambda o, la=la, lb=lb: {
                "keine": "keine zwingende Abhängigkeit",
                "a_vor_b": f"„{la}“ muss vor „{lb}“ liegen",
                "b_vor_a": f"„{lb}“ muss vor „{la}“ liegen"}[o])
        if wahl != status:
            core.clear_abhaengigkeit(state, a["id"], b["id"])
            if wahl == "a_vor_b":
                core.set_abhaengigkeit(state, a["id"], b["id"])
            elif wahl == "b_vor_a":
                core.set_abhaengigkeit(state, b["id"], a["id"])
            st.rerun()


def _einheiten_portfolio(state, einheiten, types, features, weights):
    """Streudiagramm der Einheiten: Gesamtaufwand gegen erreichbaren
    Aehnlichkeitsgewinn. Zeigt, welche Einheit viel Wirkung fuer wenig Aufwand
    bietet, und stuetzt die Priorisierung visuell."""
    import pandas as pd
    df = pd.DataFrame([{
        "Betrachtungseinheit": core.get_name(state, u),
        "Aufwand": core.einheit_kennzahlen(state, u, types, features, weights)["aufwand"],
        "Gewinn": core.einheit_kennzahlen(state, u, types, features, weights)["gewinn"],
    } for u in einheiten])
    try:
        import altair as alt
        basis = alt.Chart(df).encode(
            x=alt.X("Aufwand:Q",
                    axis=alt.Axis(title="Gesamtaufwand (Summe der Aufwandsstufen)")),
            y=alt.Y("Gewinn:Q", scale=alt.Scale(domain=[0, 1]),
                    axis=alt.Axis(format="%", title="erreichbarer Ähnlichkeitsgewinn")))
        punkte = basis.mark_circle(size=220).encode(
            tooltip=["Betrachtungseinheit:N", "Aufwand:Q",
                     alt.Tooltip("Gewinn:Q", format=".0%")])
        beschriftung = basis.mark_text(dy=-14).encode(text="Betrachtungseinheit:N")
        chart = (punkte + beschriftung).properties(height=320)
        try:
            st.altair_chart(chart, width="stretch")
        except TypeError:
            st.altair_chart(chart, use_container_width=True)
    except Exception:
        st.dataframe(df, hide_index=True, width="stretch")



def _nutzwertanalyse(state, einheiten, types, features, weights):
    """Nutzwertanalyse nach Zangemeister: gewichtete, gepolte Kriterien ergeben je
    Einheit einen Nutzwert und damit die Reihenfolge, welche Einheit man zuerst
    angeht. Aufwand und Kosten sind aus den Massnahmen vorbelegt, Risiko und
    Stueckzahl gibt der Anwender ein. Kriterien sind frei ergaenzbar."""
    import pandas as pd
    st.subheader("Nutzwertanalyse")
    krit = core.nwa_kriterien(state)

    st.markdown("**Kriterien und Gewichte**")
    for k in list(krit):
        c1, c2, c3 = st.columns([3, 2, 1])
        richtung_txt = "klein ist besser" if k["richtung"] == "min" else "groß ist besser"
        c1.markdown(f"**{k['name']}**  \n{richtung_txt}")
        g = c2.number_input("Gewicht", min_value=0.0, value=float(k["gewicht"]),
                            step=5.0, key=f"nwa_g_{k['name']}",
                            label_visibility="collapsed")
        if g != k["gewicht"]:
            core.nwa_set_gewicht(state, k["name"], g)
            st.rerun()
        if c3.button("Entfernen", key=f"nwa_del_{k['name']}"):
            core.nwa_remove_kriterium(state, k["name"])
            st.rerun()
    summe = sum(k["gewicht"] for k in krit)
    st.caption(f"Summe der Gewichte: {summe:.0f}. Sie werden intern auf 100 % "
               "normiert, müssen sich also nicht exakt zu 100 addieren.")

    c1, c2, c3 = st.columns([3, 2, 1])
    neu_name = c1.text_input("Neues Kriterium", key="nwa_neu",
                             label_visibility="collapsed",
                             placeholder="Neues Kriterium")
    neu_r = c2.selectbox("Richtung", ["min", "max"], key="nwa_neu_r",
                         format_func=lambda r: "klein ist besser" if r == "min"
                         else "groß ist besser", label_visibility="collapsed")
    if c3.button("Hinzufügen", key="nwa_add") and neu_name.strip():
        core.nwa_add_kriterium(state, neu_name, neu_r)
        st.rerun()

    st.markdown("**Bewertung je Betrachtungseinheit**")
    # Spaltenueberschrift kennzeichnet, welche Werte schon aus den Aenderungen
    # vorliegen. Der Kriterienname selbst bleibt davon unberuehrt.
    spalte = {k["name"]: (f"{k['name']} (bereits erfasst)" if k["auto"]
                          else k["name"]) for k in krit}
    daten = []
    for u in einheiten:
        zeile = {"Betrachtungseinheit": core.get_name(state, u)}
        for k in krit:
            zeile[spalte[k["name"]]] = core.nwa_get_wert(
                state, u, k, types, features, weights)
        daten.append(zeile)
    df = pd.DataFrame(daten)
    cfg = {"Betrachtungseinheit": st.column_config.TextColumn(disabled=True)}
    for k in krit:
        cfg[spalte[k["name"]]] = st.column_config.NumberColumn()
    edit = st.data_editor(df, hide_index=True, width="stretch",
                          key=f"nwa_editor_{len(krit)}_{len(einheiten)}",
                          column_config=cfg)
    for i, u in enumerate(einheiten):
        kz = core.einheit_kennzahlen(state, u, types, features, weights)
        for k in krit:
            wert = edit.iloc[i][spalte[k["name"]]]
            wert = float(wert) if pd.notna(wert) else 0.0
            if k["auto"] and abs(wert - kz[k["auto"]]) < 1e-9:
                core.nwa_set_wert(state, u, k["name"], None)
            else:
                core.nwa_set_wert(state, u, k["name"], wert)

    unvollstaendig = []
    for u in einheiten:
        v = core.einheit_kennzahlen(
            state, u, types, features, weights)["kosten_vollstaendig"]
        if v < 1.0:
            unvollstaendig.append(f"{core.get_name(state, u)} ({v * 100:.0f} %)")
    if unvollstaendig:
        st.caption("Hinweis zu den Kosten: Nicht alle Änderungen dieser "
                   "Betrachtungseinheiten "
                   "sind beziffert, die Kosten sind aus dem Durchschnitt "
                   "hochgerechnet (Anteil bezifferter Änderungen in Klammern): "
                   + ", ".join(unvollstaendig) + ". Sie können die Werte oben "
                   "jederzeit überschreiben.")

    st.markdown("**Rangfolge nach Nutzwert**")
    nw = core.nwa_nutzwerte(state, types, features, weights)
    erg = pd.DataFrame([{"Betrachtungseinheit": core.get_name(state, u),
                         "Nutzwert": v} for u, v in nw])
    try:
        import altair as alt
        chart = (alt.Chart(erg).mark_bar().encode(
            x=alt.X("Nutzwert:Q", scale=alt.Scale(domain=[0, 1]),
                    axis=alt.Axis(format="%")),
            y=alt.Y("Betrachtungseinheit:N", sort="-x", title=None),
            tooltip=["Betrachtungseinheit:N",
                     alt.Tooltip("Nutzwert:Q", format=".1%")])
            .properties(height=60 + 40 * len(erg)))
        try:
            st.altair_chart(chart, width="stretch")
        except TypeError:
            st.altair_chart(chart, use_container_width=True)
    except Exception:
        st.dataframe(erg, hide_index=True, width="stretch")
    if nw:
        st.markdown(f"Zuerst angehen: **{core.get_name(state, nw[0][0])}**.")


def _priorisierung(state, features, types, weights):
    """Vierter Schritt, Teil 2: Priorisierung der Betrachtungseinheiten, also welche
    Einheit man zuerst angeht. Ein Portfolio stellt Aufwand und Wirkung je Einheit
    gegenueber, die Nutzwertanalyse liefert die begruendete Rangfolge."""
    einheiten = [u for u in state["units"] if core.get_zieltyp(state, u)]
    if not einheiten:
        return
    st.header("Priorisierung")
    st.subheader("Portfolio")
    _einheiten_portfolio(state, einheiten, types, features, weights)
    st.divider()
    _nutzwertanalyse(state, einheiten, types, features, weights)


@st.dialog("Gemeinsame Werte der inhaltlichen Synergie")
def _synergie_dialog(state, felder_menge, massnahmen):
    """Erfasst beim Zusammenlegen verschiedener Massnahmen zu einer gemeinsamen
    Massnahme die gemeinsamen Werte. Der Aehnlichkeitsgewinn addiert sich
    automatisch, hier geht es nur um den gemeinsamen Aufwand sowie die gemeinsamen
    Kosten und die Dauer. Vorschlag ist der hoechste Einzelwert."""
    von = {m["id"]: m for m in massnahmen}
    merkmale = []
    for k in felder_menge:
        if k in von:
            merkmale.extend(von[k]["merkmale"])
    st.markdown("**" + " + ".join(merkmale) + "**")
    au, ks, ds = [], [], []
    for k in felder_menge:
        if k in von:
            m = von[k]
            if m["aufwand"] in core.AUFWAND_STUFEN:
                au.append(m["aufwand"])
            if m["kosten"] is not None:
                ks.append(m["kosten"])
            if m["dauer"] is not None:
                ds.append(m["dauer"])
    st.caption("Vorschlag ist der höchste Einzelwert der beteiligten Änderungen.")
    stufen = list(core.AUFWAND_STUFEN)
    v_auf = max(au) if au else None
    idx = stufen.index(v_auf) if v_auf in stufen else None
    aufwand = st.radio("Gemeinsamer Aufwand", stufen,
                       format_func=lambda x: core.AUFWAND_LABEL[x],
                       index=idx, horizontal=True)
    kosten = st.number_input("Gemeinsame Kosten in Euro (optional)", min_value=0.0,
                             value=(max(ks) if ks else None), step=1000.0,
                             format="%.0f")
    dauer = st.number_input("Gemeinsame Dauer in Wochen (optional)", min_value=0.0,
                            value=(max(ds) if ds else None), step=1.0, format="%.0f")
    sp1, sp2 = st.columns(2)
    if sp1.button("Speichern", type="primary", width="stretch"):
        core.add_synergie(state, felder_menge, aufwand, kosten, dauer)
        st.rerun()
    if sp2.button("Abbrechen", width="stretch"):
        st.rerun()


HILFE_KONSOLIDIERUNG = """
**Schritt 4 von 4 · Konsolidierung**

Die Änderungen aller Betrachtungseinheiten werden hier unternehmensweit
zusammengeführt, auf Synergien geprüft und die Betrachtungseinheiten in eine
Reihenfolge gebracht. Grundlage sind die im vorherigen Schritt festgelegten
Zieltypen.

**Aufbau der Seite**

**1. Änderungsliste.** Alle Änderungen aller Betrachtungseinheiten in einer
fortlaufend nummerierten Tabelle.

*Änderung: ein Merkmal, dessen Ziel-Ausprägung von der aktuellen abweicht und
somit einen Handlungsbedarf für das Unternehmen darstellt.*

*Voraussetzung: was vorliegen muss, damit die Änderung umgesetzt werden kann. Diese
Spalte tragen Sie selbst ein. Bsp: ein ERP-Modul, eine Fläche am Standort oder eine
Mitarbeiterqualifizierung.*

**2. Harmonisierung.** Legen Sie Änderungen zusammen, die sich gemeinsam umsetzen
lassen. Wählen Sie sie im Auswahlfeld unter der Tabelle aus und erfassen Sie im
Fenster die gemeinsamen Werte. Dafür gibt es zwei Anlässe.

>>*Gleiches Ziel: Mehrere Betrachtungseinheiten streben bei demselben Merkmal
dieselbe Ausprägung an. Solche Fälle schlägt das Werkzeug unter der Tabelle vor.*

>>*Gleiche Voraussetzung: Verschiedene Änderungen benötigen dieselbe
Voraussetzung. Diese erkennen nur Sie, da sie sich aus Ihrer
Voraussetzungs-Spalte ergeben.*

In beiden Fällen addiert sich der Ähnlichkeitsgewinn, während Aufwand, Kosten und
Dauer nur einmal anfallen und deshalb gemeinsam neu zu schätzen sind. Genau darin
liegt der Nutzen der Harmonisierung.

**3. Priorisierung.** Zwei Werkzeuge helfen bei der Frage, welche
Betrachtungseinheit zuerst anzugehen ist.

>>*Portfolio: stellt je Betrachtungseinheit den Gesamtaufwand dem erreichbaren
Ähnlichkeitsgewinn gegenüber. Oben links liegt, wer mit wenig Aufwand viel
gewinnt.*

>>*Nutzwertanalyse: bewertet die Betrachtungseinheiten anhand gewichteter Kriterien
und ergibt eine Rangfolge. Aufwand und Kosten sind aus den Änderungen vorbelegt und
überschreibbar, die übrigen Kriterien schätzen Sie selbst. Kriterien lassen sich
ergänzen und entfernen, die Gewichte sind frei wählbar.*

**4. Ergebnis exportieren.** Die Datei enthält die Zieltypen, die Rangfolge, die
harmonisierte Änderungsliste und einen gemeinsamen morphologischen Kasten mit allen
Betrachtungseinheiten. Sie lässt sich im Browser öffnen und dort als PDF speichern.

Damit endet das Konfigurationsvorgehen. Die weitere Ausgestaltung der Änderungen
und ihre zeitliche Planung schließen sich als eigenes Vorhaben an.
"""


def seite_massnahmen(state, features, types, weights):
    st.title("Konsolidierung")

    if not state["units"]:
        st.caption("Keine Betrachtungseinheiten vorhanden.")
        return

    mit_ziel = [u for u in state["units"] if core.get_zieltyp(state, u)]
    if not mit_ziel:
        st.caption("Die Konsolidierung erscheint, sobald mindestens eine "
                   "Betrachtungseinheit einen Zieltyp hat.")
    else:
        # Reihenfolge: 1. Harmonisierung, 2. Detaillierung, 3. Priorisierung.
        _harmonisierung(state, features, types, weights)
        st.divider()
        _priorisierung(state, features, types, weights)
        st.divider()

    st.subheader("Ergebnis exportieren")
    st.download_button(
        "Gesamtergebnis als HTML herunterladen",
        data=_export_html(state, features, types, weights),
        file_name="konfigurationsergebnis.html", mime="text/html")
    st.caption("Die Datei lässt sich im Browser öffnen und über dessen "
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
        help="Datei mit Merkmalen, Ausprägungen, Typprofilen und (optional) Gewichten."
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
        st.error("Die Gewichtsspalte enthält fehlerhafte Werte:")
        for merkmal, grund in probleme:
            st.write(f"- **{merkmal}**: {grund}")
        st.write("Sie können die Excel-Datei korrigieren und neu laden, "
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
