"""
Baustein 1: Datenschicht + Zustandslogik (+ spaeter Berechnung).

Bewusst OHNE Streamlit. Alle Funktionen sind rein und testbar.
Die Oberflaeche (spaeter) haelt den Zustand nur im Session State,
ruft aber genau diese Funktionen auf.

Begriffe:
  - Merkmal             : eine Gruppe von Zeilen der Morphologie
  - Auspraegung         : eine moegliche Auswahl zu einem Merkmal
  - Betrachtungseinheit : eine vom Nutzer erzeugte Spalte (A, B, C, ...);
                          eine vollstaendige Konfiguration ueber alle Merkmale
  - OFFEN               : noch keine Ist-Angabe gesetzt (fehlende Angabe)
  - Gewicht             : relative Wichtigkeit eines Merkmals (aus der Excel)
"""

# --- Imports ---
# openpyxl: Bibliothek, die Excel-Dateien lesen kann.
import openpyxl
# OrderedDict: Dictionary, bei dem die Einfuege-Reihenfolge der Schluessel
# erhalten bleibt. Wichtig, weil die Reihenfolge der Merkmale aus der Excel
# Bedeutung traegt und nicht durcheinandergeraten darf.
from collections import OrderedDict

# --- Konstanten, die wir an mehreren Stellen brauchen ---
# OFFEN: fuer dieses Merkmal wurde (noch) keine Ist-Auspraegung gesetzt.
# Wir nehmen None, weil es sich nie mit einem echten Auswahltext verwechseln laesst.
# In der gemeinsamen Erfassung ergibt sich OFFEN automatisch, wenn der Nutzer
# bei einem Merkmal keine Auspraegung auf "Ist" stellt.
OFFEN = None

# Soll-Sentinel fuer die Detaillierung: In den Faellen 4 und 5 (kein Ist) kann
# der Nutzer bewusst NICHTS anstreben. Das ist eine getroffene Entscheidung
# (Pflicht erfuellt) und zaehlt im Soll-Score als Nichttreffer. Es ist damit
# verschieden von OFFEN, das "noch nicht entschieden" bedeutet.
NICHTS_ANSTREBEN = "\x00nichts_anstreben"


# Eigene Fehlerklasse fuer Probleme beim Einlesen der Gewichte.
# So kann die Oberflaeche gezielt darauf reagieren (Meldung + Wahl anbieten).
class GewichtFehler(Exception):
    """Wird ausgeloest, wenn die Gewichtsspalte existiert, aber fehlerhafte
    Werte enthaelt (fehlend, nicht-numerisch, oder <= 0).

    Traegt die schon eingelesene Morphologie mit, damit die Oberflaeche im
    Fehlerfall trotzdem 'ungewichtet fortfahren' anbieten kann:
      .probleme : Liste (merkmal, grund)
      .features : die eingelesenen Merkmale/Auspraegungen
      .types    : die eingelesenen Typprofile
    """
    def __init__(self, probleme, features=None, types=None):
        self.probleme = probleme
        self.features = features
        self.types = types
        text = "; ".join(f"{m}: {g}" for m, g in probleme)
        super().__init__(f"Fehlerhafte Gewichte: {text}")


# ============================================================ DATENSCHICHT
def _gewicht_zu_zahl(roh):
    """Wandelt einen rohen Gewichtswert aus der Excel in eine Zahl (float) um.
    Akzeptiert sowohl echte Zahlen als auch Texte mit Komma ODER Punkt als
    Dezimaltrennzeichen (z.B. '0,5' und '0.5' ergeben beide 0.5).
    Gibt None zurueck, wenn keine Umwandlung moeglich ist."""
    # Schon eine Zahl? Direkt uebernehmen.
    if isinstance(roh, (int, float)):
        return float(roh)
    # Text: Komma zu Punkt, Leerzeichen entfernen, dann umwandeln versuchen.
    if isinstance(roh, str):
        text = roh.strip().replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return None
    # Alles andere (z.B. None) ist keine gueltige Zahl.
    return None


def load_morphology(path_or_buffer):
    """Liest die Excel-Morphologie im NEUEN Layout.

    path_or_buffer: Dateipfad ODER bereits geoeffnete Datei (z.B. Streamlit-
    Hochladefeld). Beides funktioniert mit derselben Funktion.

    Erwartetes Layout:
      Zeile 1: 'Typen' | (leer) | (leer) | Typ-Klarnamen ab Spalte D
      Zeile 2: 'Merkmal' | 'Gewicht' | 'Auspraegung' | Typ 1 | Typ 2 | ...
      ab Zeile 3: Merkmal (NUR in erster Zeile je Merkmal) | Gewicht (nur 1x)
                  | Auspraegung | x/leer je Typ
    Der Merkmalsname wird per "Auffuellen nach unten" (forward fill) auf die
    folgenden Zeilen uebertragen, bis ein neuer Merkmalsname auftaucht.
    Das Gewicht steht ebenfalls nur in der ersten Zeile eines Merkmals.

    Die Gewichtsspalte ist OPTIONAL:
      - fehlt sie ganz  -> weights = None (Oberflaeche rechnet gleich gewichtet)
      - existiert sie   -> Werte werden geprueft; bei Fehlern: GewichtFehler

    Rueckgabe:
      features : OrderedDict  merkmal -> [auspraegung, ...]
      types    : dict         typ_klarname -> { merkmal -> set(zulaessige auspraegungen) }
      weights  : dict|None    merkmal -> float (>0)  ODER None (keine Spalte)
    """
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    klar = rows[0]      # Zeile 1: Typ-Klarnamen (ab Spalte D = Index 3)
    header = rows[1]    # Zeile 2: Spaltenkoepfe

    # Pruefen, ob es eine Gewichtsspalte gibt (Spalte B / Index 1 == "Gewicht").
    hat_gewichtsspalte = (len(header) > 1 and header[1] is not None
                          and str(header[1]).strip().lower() == "gewicht")

    # Spaltenpositionen je nach Layout.
    # Mit Gewicht:  A=Merkmal(0) B=Gewicht(1) C=Auspraegung(2) D..=Typen(3..)
    # Ohne Gewicht: A=Merkmal(0) B=Auspraegung(1) C..=Typen(2..)
    if hat_gewichtsspalte:
        col_merkmal, col_gewicht, col_auspr, col_typ_start = 0, 1, 2, 3
    else:
        col_merkmal, col_gewicht, col_auspr, col_typ_start = 0, None, 1, 2

    # Typ-Spalten erfassen: Spaltenindex -> Klarname.
    type_cols = OrderedDict()
    for col in range(col_typ_start, len(header)):
        if header[col] is not None:
            name = klar[col] if (col < len(klar) and klar[col]) else header[col]
            type_cols[col] = str(name).strip()

    features = OrderedDict()
    profiles = {name: {} for name in type_cols.values()}
    roh_gewichte = {}   # merkmal -> roher Zellwert aus der ersten Zeile

    aktuelles_merkmal = None   # fuer das Auffuellen nach unten

    for r in rows[2:]:
        if not r:
            continue
        # Merkmalszelle lesen. Ist sie gefuellt, beginnt ein neues Merkmal.
        zell_merkmal = r[col_merkmal] if col_merkmal < len(r) else None
        if zell_merkmal is not None and str(zell_merkmal).strip() != "":
            aktuelles_merkmal = str(zell_merkmal).strip()
            # Gewicht NUR aus dieser ersten Zeile des Merkmals lesen.
            if col_gewicht is not None:
                roh = r[col_gewicht] if col_gewicht < len(r) else None
                roh_gewichte[aktuelles_merkmal] = roh
        # Wenn noch kein Merkmal bekannt ist, Zeile ueberspringen.
        if aktuelles_merkmal is None:
            continue

        # Auspraegung lesen.
        auspr = r[col_auspr] if col_auspr < len(r) else None
        if auspr is None or str(auspr).strip() == "":
            continue
        auspr = str(auspr).strip()

        features.setdefault(aktuelles_merkmal, [])
        if auspr not in features[aktuelles_merkmal]:
            features[aktuelles_merkmal].append(auspr)

        for col, name in type_cols.items():
            val = r[col] if col < len(r) else None
            if val is not None and str(val).strip().lower() == "x":
                profiles[name].setdefault(aktuelles_merkmal, set()).add(auspr)

    # --- Gewichte aufbereiten und (streng) pruefen ---
    if not hat_gewichtsspalte:
        weights = None
    else:
        weights = {}
        probleme = []   # (merkmal, grund)
        for merkmal in features:
            roh = roh_gewichte.get(merkmal, None)
            if roh is None or (isinstance(roh, str) and roh.strip() == ""):
                probleme.append((merkmal, "Gewicht fehlt"))
                continue
            wert = _gewicht_zu_zahl(roh)
            if wert is None:
                probleme.append((merkmal, f"Gewicht ist keine Zahl ({roh!r})"))
                continue
            if wert <= 0:
                probleme.append((merkmal, f"Gewicht muss > 0 sein ({wert})"))
                continue
            weights[merkmal] = wert
        if probleme:
            # Streng: bei JEDEM Problem abbrechen, damit die Oberflaeche
            # die Wahl "korrigieren ODER ungewichtet fortfahren" anbieten kann.
            # features/types werden mitgegeben, damit 'ungewichtet fortfahren'
            # ohne erneutes Einlesen moeglich ist.
            raise GewichtFehler(probleme, features=features, types=profiles)

    return features, profiles, weights


# ============================================================ ZUSTANDSLOGIK
def empty_state():
    """Leerer Anfangszustand: noch keine Betrachtungseinheiten.
    Der Nutzer erzeugt die erste Einheit selbst."""
    return {
        # Liste der Einheiten-IDs in Anlege-Reihenfolge, z.B. ["A", "B"].
        "units": [],
        # IST-Auswahl je Einheit: unit_id -> { merkmal -> auspraegung }.
        # Genau EINE Auspraegung pro Merkmal (homogen).
        "matrix": {},
        # POTENTIAL je Einheit: unit_id -> { merkmal -> set(auspraegungen) }.
        # MEHRERE moeglich; beschreibt den Moeglichkeitsraum (Phase 3).
        "potential": {},
        # AUSSCHLUSS je Einheit: unit_id -> { merkmal -> set(auspraegungen) }.
        # MEHRERE moeglich; grundsaetzlich ausgeschlossene Auspraegungen.
        "ausschluss": {},
        # ENGERE AUSWAHL je Einheit: unit_id -> [typ_name, ...] (max. 3).
        # Die Typen, die der Nutzer zum Vergleich in die Detaillierung nimmt.
        # Reihenfolge = Auswahl-Reihenfolge (bestimmt die Tab-Reihenfolge).
        "engere_auswahl": {},
        # ZIELTYP je Einheit: unit_id -> typ_name ODER None (noch nicht festgelegt).
        # Der FINALE Zieltyp, den der Nutzer nach dem Vergleich aus der engeren
        # Auswahl bestaetigt. Bis dahin None; die Tabs sind reiner Vergleich.
        "zieltyp": {},
        # SOLL-Auspraegung je Einheit UND Typ (Schritt 6, Detaillierung):
        # unit_id -> { typ_name -> { merkmal -> angestrebte auspraegung ODER OFFEN } }.
        # Jeder Typ der engeren Auswahl hat eine EIGENE Soll-Konfiguration, damit
        # sich die Typen unabhaengig voneinander detaillieren und vergleichen lassen.
        "soll": {},
        # AUFWAND je Einheit, Typ und Merkmal (Schritt 6, Detaillierung):
        # unit_id -> { typ_name -> { merkmal -> 1|2|3 } }. Pflichtangabe, sobald
        # ein Soll gewaehlt ist, das vom Ist abweicht. 1=gering, 2=mittel, 3=hoch.
        # Fehlt fuer ein Merkmal ein Eintrag, ist der Aufwand noch nicht geschaetzt.
        "aufwand": {},
        # KOSTEN je (uid, typ, merkmal): optionale monetaere Schaetzung in Euro
        # (float) oder None. Analog zu aufwand, aber freiwillig. Grundlage der
        # spaeteren Budgetkurve; fehlende Werte bleiben None.
        "kosten": {},
        # DAUER je (uid, typ, merkmal): optionale Schaetzung der Umsetzungsdauer
        # in Wochen (float) oder None. Analog zu kosten, freiwillig.
        "dauer": {},
        # WELT je Einheit und Merkmal (Tims Regel, explizit ueber Dropdown):
        # unit_id -> { merkmal -> "ist_bekannt" | "ist_unbekannt" }. Standard: ist_bekannt.
        "welt": {},
        # GEWICHTE je Merkmal (global, nicht je Einheit): merkmal -> float (>= 1).
        # Relative Wichtigkeit fuer die Aehnlichkeitsmasse; 1 = normal, 2 = doppelt,
        # usw. Wird beim Laden aus den Excel-Startwerten (falls vorhanden) oder mit
        # 1 initialisiert und in der Zustandserfassung editiert.
        "gewichte": {},
        # MASSNAHMEN je Einheit und Merkmal (nur fuer Handlungsfelder des Zieltyps):
        # unit_id -> { merkmal -> {"text": str, "phase": 1|2|3, "wer": str} }.
        # Der Freitext und die Verantwortlichkeit stammen vollstaendig vom Anwender,
        # die Phase ist aus dem Aufwand vorbelegt und aenderbar.
        "massnahmen": {},
        # BUENDEL: Menge von (merkmal, soll)-Paaren, die der Anwender ueber die
        # Betrachtungseinheiten hinweg zu einem gemeinsamen Handlungsfeld
        # zusammengelegt hat. Nur Paare, die bei mehreren Einheiten als Handlungsfeld
        # auftreten, sind buendelbar.
        "buendel": set(),
        # BUENDEL_WERTE je zusammengelegter Kombination (merkmal, soll): gemeinsamer
        # Aufwand, Kosten und Dauer, vom Anwender im Zusammenlege-Dialog erfasst.
        # dict {(merkmal, soll): {"aufwand": int|None, "kosten": float|None,
        # "dauer": float|None}}.
        "buendel_werte": {},
        # SYNERGIEN: frei gebildete inhaltliche Zusammenlegungen verschiedener
        # Handlungsfelder zu einer gemeinsamen Massnahme. Liste von dicts mit
        # "felder" (frozenset von hf_schluessel), "aufwand", "kosten", "dauer".
        "synergien": [],
        # ABHAENGIGKEITEN: zwingende Reihenfolge zwischen Massnahmen als Menge von
        # (vorher_id, nachher_id). Die erste Massnahme muss vor der zweiten liegen.
        "abhaengigkeiten": set(),
        # HF_MASSNAHMEN: Massnahmen der unternehmensweiten Handlungsfelder, je
        # Handlungsfeld ueber einen Schluessel (siehe hf_schluessel) statt je Einheit.
        "hf_massnahmen": {},
        # NAMEN je Einheit: unit_id -> frei vergebener Anzeigename (String).
        # Rein sichtbar; die interne unit_id (Buchstabe) bleibt als Schluessel
        # stabil. Leer = kein Name gesetzt, dann wird ersatzweise die unit_id
        # angezeigt. Namen duerfen sich doppeln, ohne dass etwas bricht.
        "namen": {},
        # Phase: "konfiguration" | "ergebnis" | "potential" | "ergebnis_soll"
        #        | "detaillierung".
        "phase": "konfiguration",
        # Zaehlt, wie viele Einheiten JEMALS vergeben wurden. Zaehlt nur hoch,
        # nie zurueck -> ein Buchstabe wird nie wiederverwendet (eindeutig ueber Zeit).
        "vergeben": 0,
    }


def _letter_at(index):
    """Buchstabe zur laufenden Nummer (0->A, 1->B, ..., 25->Z, 26->AA, ...).
    Der fuehrende Unterstrich signalisiert: internes Hilfsmittel."""
    n = index + 1
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def add_unit(state, features):
    """Legt eine neue Betrachtungseinheit an. Alle Ist-Zellen = OFFEN
    (noch nicht angeschaut), damit der Nutzer jedes Merkmal bewusst
    beantworten muss, bevor ausgewertet werden kann.
    Potential und Ausschluss starten leer (pro Merkmal eine leere Menge).
    Der Buchstabe richtet sich nach dem Zaehler 'vergeben' -> nie Recycling.
    Gibt die neue unit_id zurueck."""
    uid = _letter_at(state["vergeben"])
    state["vergeben"] += 1
    state["units"].append(uid)
    state["matrix"][uid] = {m: OFFEN for m in features}
    state["potential"][uid] = {m: set() for m in features}
    state["ausschluss"][uid] = {m: set() for m in features}
    state["engere_auswahl"][uid] = []
    state["zieltyp"][uid] = None
    state["soll"][uid] = {}   # je Typ eine eigene Soll-Konfiguration (lazy angelegt)
    state["aufwand"][uid] = {}   # je Typ die Aufwaende der Merkmale (lazy angelegt)
    state["kosten"][uid] = {}    # je Typ die optionalen Kosten der Merkmale
    state["dauer"][uid] = {}     # je Typ die optionalen Dauern der Merkmale
    state["massnahmen"][uid] = {}   # Massnahmen je Merkmal (lazy angelegt)
    state["welt"][uid] = {m: WELT_IST_BEKANNT for m in features}
    state["namen"][uid] = ""   # Anzeigename, vom Anwender frei vergebbar
    return uid


def remove_unit(state, uid):
    """Entfernt eine Betrachtungseinheit vollstaendig (Ist, Potential,
    Ausschluss, engere Auswahl, Zieltyp, Ziel, Aufwand, Welt). Der Zaehler
    'vergeben' bleibt unveraendert."""
    for ebene in ("matrix", "potential", "ausschluss", "engere_auswahl",
                  "zieltyp", "soll", "aufwand", "kosten", "dauer", "welt", "massnahmen", "namen"):
        if uid in state[ebene]:
            del state[ebene][uid]
    if uid in state["units"]:
        state["units"].remove(uid)


def get_name(state, uid):
    """Anzeigename einer Einheit. Der frei vergebene Name, falls gesetzt,
    sonst die interne unit_id (Buchstabe). So bleibt jede Einheit auch ohne
    Namen eindeutig benannt."""
    name = state.get("namen", {}).get(uid, "")
    return name if name else uid


def set_name(state, uid, name):
    """Setzt den Anzeigenamen einer Einheit. Fuehrende und schliessende
    Leerzeichen werden entfernt; ein leerer Name faellt auf die unit_id
    zurueck (siehe get_name). Beruehrt keine Logik, nur die Anzeige."""
    state.setdefault("namen", {})[uid] = (name or "").strip()


def set_choice(state, uid, merkmal, auspraegung):
    """Setzt fuer eine Einheit bei einem Merkmal genau EINE Auspraegung.
    Ueberschreibt die vorherige -> nie mehr als eine Auswahl pro Zelle."""
    state["matrix"][uid][merkmal] = auspraegung


def get_choice(state, uid, merkmal):
    """Liefert die aktuelle Ist-Auswahl. OFFEN bedeutet 'noch nicht angeschaut'."""
    return state["matrix"][uid].get(merkmal, OFFEN)


# --- Potential / Ausschluss (Phase 3) ---
# Pro Einheit, Merkmal und Auspraegung gibt es genau einen von drei Zustaenden:
#   "" (nichts) | "potential" | "ausschluss"
# Da eine Auspraegung immer nur EINEN dieser Zustaende hat, kann sie nie
# gleichzeitig Potential und Ausschluss sein (Regel ist selbsttragend).

def set_pa_status(state, uid, merkmal, auspraegung, status):
    """Setzt den Potential/Ausschluss-Vergleichsfall EINER Auspraegung.
    status: "" -> weder/noch, "potential" -> Potential, "ausschluss" -> Ausschluss.
    Entfernt die Auspraegung zuerst aus beiden Mengen und fuegt sie dann
    in die passende ein. So ist garantiert: nie in beiden gleichzeitig."""
    state["potential"][uid][merkmal].discard(auspraegung)
    state["ausschluss"][uid][merkmal].discard(auspraegung)
    if status == "potential":
        state["potential"][uid][merkmal].add(auspraegung)
    elif status == "ausschluss":
        state["ausschluss"][uid][merkmal].add(auspraegung)


def get_pa_status(state, uid, merkmal, auspraegung):
    """Liefert den Vergleichsfall EINER Auspraegung: "", "potential" oder "ausschluss"."""
    if auspraegung in state["potential"][uid][merkmal]:
        return "potential"
    if auspraegung in state["ausschluss"][uid][merkmal]:
        return "ausschluss"
    return ""


def get_potential(state, uid, merkmal):
    """Menge der als Potential markierten Auspraegungen (kann leer sein)."""
    return state["potential"][uid][merkmal]


def get_ausschluss(state, uid, merkmal):
    """Menge der als Ausschluss markierten Auspraegungen (kann leer sein)."""
    return state["ausschluss"][uid][merkmal]


def passende_potenziale(state, uid, merkmal, typ_profil):
    """Liefert die Potenzial-Auspraegungen dieses Merkmals, die im Profil des
    Zieltyps liegen (also den Typ tatsaechlich erreichen). Sortierte Liste.
    Grundlage fuer die Detaillierung: bei genau einem passenden Potenzial wird
    es automatisch angestrebt, bei mehreren muss der Nutzer waehlen."""
    profil = typ_profil.get(merkmal, set())
    return sorted(state["potential"][uid][merkmal] & profil)


# --- Kombinierter Zellzustand fuer die gemeinsame Erfassung (durchschaltbar) ---
# Eine Auspraegung hat genau EINEN von vier Zustaenden:
ZELLE_OFFEN = "offen"          # nichts gesetzt
ZELLE_IST = "ist"             # die (eine) Ist-Auspraegung des Merkmals
ZELLE_POTENTIAL = "potential"  # als Potenzial markiert
ZELLE_AUSSCHLUSS = "ausschluss"  # als Ausschluss markiert

# Reihenfolge beim Durchklicken: offen -> ist -> potential -> ausschluss -> offen
_ZELLE_REIHENFOLGE = [ZELLE_OFFEN, ZELLE_IST, ZELLE_POTENTIAL, ZELLE_AUSSCHLUSS]

# --- Welt je Merkmal und Einheit (Tims Regel, explizit ueber Dropdown) ---
WELT_IST_BEKANNT = "ist_bekannt"       # Klicken erzeugt Ist, dann Potenziale
WELT_IST_UNBEKANNT = "ist_unbekannt"   # Klicken erzeugt nur Ausschluesse


def get_zellzustand(state, uid, merkmal, auspraegung):
    """Liefert den kombinierten Zustand EINER Auspraegung:
    ZELLE_IST, ZELLE_POTENTIAL, ZELLE_AUSSCHLUSS oder ZELLE_OFFEN.
    Beruecksichtigt alle drei Ebenen (Ist hat Vorrang)."""
    if get_choice(state, uid, merkmal) == auspraegung:
        return ZELLE_IST
    pa = get_pa_status(state, uid, merkmal, auspraegung)
    if pa == "potential":
        return ZELLE_POTENTIAL
    if pa == "ausschluss":
        return ZELLE_AUSSCHLUSS
    return ZELLE_OFFEN


def _setze_zellzustand(state, uid, merkmal, auspraegung, neuer_zustand):
    """Setzt eine Auspraegung auf einen bestimmten kombinierten Zustand und
    haelt dabei alle Ebenen konsistent:
      - ZELLE_IST: setzt diese Auspraegung als Ist (genau EINE pro Merkmal,
        die vorherige Ist-Auspraegung weicht automatisch) und entfernt sie
        aus Potential/Ausschluss.
      - ZELLE_POTENTIAL / ZELLE_AUSSCHLUSS: setzt den PA-Vergleichsfall; falls die
        Auspraegung bisher das Ist war, wird das Ist auf OFFEN gesetzt.
      - ZELLE_OFFEN: entfernt sie aus Potential/Ausschluss; war sie das Ist,
        wird das Ist auf OFFEN gesetzt."""
    war_ist = (get_choice(state, uid, merkmal) == auspraegung)

    if neuer_zustand == ZELLE_IST:
        # Aus PA entfernen und als (einzige) Ist-Auspraegung setzen.
        set_pa_status(state, uid, merkmal, auspraegung, "")
        set_choice(state, uid, merkmal, auspraegung)  # ueberschreibt altes Ist
    elif neuer_zustand == ZELLE_POTENTIAL:
        if war_ist:
            set_choice(state, uid, merkmal, OFFEN)
        set_pa_status(state, uid, merkmal, auspraegung, "potential")
    elif neuer_zustand == ZELLE_AUSSCHLUSS:
        if war_ist:
            set_choice(state, uid, merkmal, OFFEN)
        set_pa_status(state, uid, merkmal, auspraegung, "ausschluss")
    else:  # ZELLE_OFFEN
        if war_ist:
            set_choice(state, uid, merkmal, OFFEN)
        set_pa_status(state, uid, merkmal, auspraegung, "")


def _merkmal_hat_ist(state, uid, merkmal):
    """True, wenn fuer dieses Merkmal und diese Einheit bereits ein Ist gesetzt ist."""
    return get_choice(state, uid, merkmal) is not OFFEN


def get_welt(state, uid, merkmal):
    """Liefert die Welt des Merkmals fuer diese Einheit
    (WELT_IST_BEKANNT oder WELT_IST_UNBEKANNT). Standard: ist_bekannt."""
    return state["welt"].get(uid, {}).get(merkmal, WELT_IST_BEKANNT)


def set_welt(state, uid, merkmal, welt):
    """Setzt die Welt und raeumt die in der neuen Welt unzulaessigen Angaben auf.
    - Wechsel zu ist_unbekannt: Ist und alle Potenziale dieses Merkmals entfernen
      (in der Kein-Ist-Welt nicht erlaubt). Ausschluesse bleiben.
    - Wechsel zu ist_bekannt: alle Ausschluesse dieses Merkmals entfernen
      (in der Ist-Welt nicht erlaubt). Ist/Potenziale bleiben."""
    state["welt"][uid][merkmal] = welt
    if welt == WELT_IST_UNBEKANNT:
        set_choice(state, uid, merkmal, OFFEN)
        for a in list(state["potential"][uid][merkmal]):
            set_pa_status(state, uid, merkmal, a, "")
    else:  # WELT_IST_BEKANNT
        for a in list(state["ausschluss"][uid][merkmal]):
            set_pa_status(state, uid, merkmal, a, "")


def klick_erfassung(state, uid, merkmal, auspraegung):
    """Klick auf eine Auspraegung in der gemeinsamen Erfassung, abhaengig von der
    ueber das Dropdown gewaehlten Welt (Tims Regel):

    WELT_IST_UNBEKANNT:
      - schaltet zwischen offen <-> Ausschluss (beliebig viele moeglich, nie ein Ist).

    WELT_IST_BEKANNT:
      - Solange KEIN Ist im Merkmal gesetzt ist: Klick macht die Auspraegung zum Ist.
      - Sobald ein Ist existiert: Klick auf eine ANDERE Auspraegung schaltet diese
        zwischen offen <-> Potenzial.
      - Klick auf die Ist-Auspraegung selbst: setzt das Merkmal komplett zurueck
        (Ist UND alle Potenziale werden geloescht).
    """
    welt = get_welt(state, uid, merkmal)
    zustand = get_zellzustand(state, uid, merkmal, auspraegung)

    if welt == WELT_IST_UNBEKANNT:
        if zustand == ZELLE_AUSSCHLUSS:
            set_pa_status(state, uid, merkmal, auspraegung, "")
        else:
            set_pa_status(state, uid, merkmal, auspraegung, "ausschluss")
        return

    # WELT_IST_BEKANNT
    if zustand == ZELLE_IST:
        # Klick auf das Ist -> Merkmal zuruecksetzen (Ist + alle Potenziale).
        set_choice(state, uid, merkmal, OFFEN)
        for a in list(state["potential"][uid][merkmal]):
            set_pa_status(state, uid, merkmal, a, "")
        return

    if not _merkmal_hat_ist(state, uid, merkmal):
        # Noch kein Ist -> diese Auspraegung wird das Ist.
        set_pa_status(state, uid, merkmal, auspraegung, "")
        set_choice(state, uid, merkmal, auspraegung)
        return

    # Es gibt bereits ein Ist auf einer anderen Auspraegung ->
    # Klick schaltet diese Auspraegung zwischen offen <-> Potenzial.
    if zustand == ZELLE_POTENTIAL:
        set_pa_status(state, uid, merkmal, auspraegung, "")
    else:
        set_pa_status(state, uid, merkmal, auspraegung, "potential")


def erfassung_unvollstaendig(state, features):
    """Prueft die Weitergehen-Bedingung fuer die gemeinsame Erfassung.
    Gibt eine Liste von Problemen zurueck (leer = alles ok). Jedes Problem ist
    ein Tupel (uid, merkmal, grund) mit grund aus:
      - 'ist_fehlt'   : Welt ist_bekannt gewaehlt, aber kein Ist gesetzt.
    Potenziale und Ausschluesse sind optional und werden NICHT verlangt.
    (Die Welt hat immer einen Wert - Standard ist_bekannt -, daher gibt es
    keinen 'welt_fehlt'-Fall; die Sperre greift ueber fehlende Ist-Angaben.)"""
    probleme = []
    for uid in state["units"]:
        for m in features:
            if get_welt(state, uid, m) == WELT_IST_BEKANNT:
                if get_choice(state, uid, m) is OFFEN:
                    probleme.append((uid, m, "ist_fehlt"))
    return probleme


def klick_zelle(state, uid, merkmal, auspraegung):
    """VERALTET (Vier-Zustand-Durchschalten). Bleibt fuer Rueckwaertskompatibilitaet
    erhalten, wird aber nicht mehr verwendet. Neue Bedienung: klick_erfassung()
    zusammen mit dem Welt-Dropdown (set_welt/get_welt)."""
    aktuell = get_zellzustand(state, uid, merkmal, auspraegung)
    idx = _ZELLE_REIHENFOLGE.index(aktuell)
    naechster = _ZELLE_REIHENFOLGE[(idx + 1) % len(_ZELLE_REIHENFOLGE)]
    _setze_zellzustand(state, uid, merkmal, auspraegung, naechster)


def is_complete(state, uid, features):
    """True, wenn der Nutzer JEDES Merkmal bewusst beantwortet hat.
    'Beantwortet' heisst: eine Ist-Auspraegung liegt vor (nicht OFFEN).
    Nur OFFEN (noch nicht angeschaut) blockiert die Auswertung."""
    return all(state["matrix"][uid][m] is not OFFEN for m in features)


def open_count(state, uid, features):
    """Anzahl der noch nicht angeschauten (OFFEN) Merkmale dieser Einheit.
    Praktisch fuer die Oberflaeche: 'noch 3 Merkmale offen'."""
    return sum(1 for m in features if state["matrix"][uid][m] is OFFEN)


# ============================================================ BERECHNUNG / ERGEBNIS
# Reine Funktionen ohne Streamlit. Sie bekommen die Auswahl einer Einheit
# (merkmal -> auspraegung), die Typprofile und optional die Gewichte und
# liefern Kennzahlen. Sie veraendern nichts.
#
# Begriff "Treffer" pro Merkmal (Containment-Koeffizient, vereinfacht):
#   |A ∩ B| / |A|  mit A = {gewaehlte Auspraegung} (immer 1 Element)
#   -> 1, wenn die gewaehlte Auspraegung im Typ-Set B liegt, sonst 0.
#   Ein OFFENes Merkmal ist nie ein Treffer.

def _gewicht(weights, merkmal):
    """Gewicht eines Merkmals. Ohne Gewichte (None) zaehlt jedes Merkmal 1."""
    if weights is None:
        return 1.0
    return float(weights.get(merkmal, 1.0))


def get_gewicht(state, merkmal):
    """Relatives Gewicht eines Merkmals (>= 1). Default 1, falls nicht gesetzt."""
    return state.get("gewichte", {}).get(merkmal, 1.0)


def set_gewicht(state, merkmal, wert):
    """Setzt das relative Gewicht eines Merkmals. Untergrenze 1."""
    try:
        w = float(wert)
    except (TypeError, ValueError):
        w = 1.0
    state.setdefault("gewichte", {})[merkmal] = max(1.0, w)


def gewicht_init(state, features, excel_weights):
    """Initialisiert die Gewichte einmalig je Merkmal: aus den Excel-Startwerten,
    falls vorhanden und gueltig, sonst mit 1. Idempotent - bereits gesetzte (im
    Werkzeug editierte) Gewichte bleiben unveraendert."""
    gew = state.setdefault("gewichte", {})
    for m in features:
        if m in gew:
            continue
        start = 1.0
        if excel_weights is not None and excel_weights.get(m) is not None:
            try:
                start = max(1.0, float(excel_weights.get(m)))
            except (TypeError, ValueError):
                start = 1.0
        gew[m] = start


def _ist_treffer(auswahl_wert, typ_profil, merkmal):
    """True, wenn die gewaehlte Auspraegung im Zulaessigkeitsraum des Typs
    fuer dieses Merkmal liegt. OFFEN ist nie ein Treffer."""
    if auswahl_wert is OFFEN:
        return False
    return auswahl_wert in typ_profil.get(merkmal, set())


def _ist_blockiert(ausschluss_unit, typ_profil, merkmal):
    """Blockade-Indikator b_i: True, wenn ALLE im Profil zulaessigen
    Auspraegungen dieses Merkmals ausgeschlossen wurden.
    Ein blockiertes Merkmal kann den Typ nie treffen und gilt daher als
    'entschiedener Nichttreffer', der in die Betrachtungsmenge einfliesst.
    ausschluss_unit darf None sein (dann nie blockiert)."""
    if ausschluss_unit is None:
        return False
    profil = typ_profil.get(merkmal, set())
    if not profil:
        return False  # leeres Profil ist nie 'komplett ausgeschlossen'
    ausgeschlossen = ausschluss_unit.get(merkmal, set())
    return profil <= ausgeschlossen


def uebereinstimmung(auswahl, typ_profil, features, weights=None, ausschluss_unit=None):
    """HAUPT-KPI (IST-MINIMUM): gewichteter Uebereinstimmungsgrad mit EINEM Typ.

    auswahl        : dict  merkmal -> gewaehlte auspraegung
    typ_profil     : dict  merkmal -> set zulaessiger auspraegungen
    features       : Reihenfolge/Menge der Merkmale
    weights        : dict merkmal->Zahl ODER None (dann alle gleich)
    ausschluss_unit: dict merkmal->set ODER None. Wenn gesetzt, gelten Merkmale
                     mit komplett ausgeschlossenem Profil als Nichttreffer im
                     vollen Nenner (Blockade-Indikator). Fuer das Minimum aendert
                     das den Wert faktisch nicht (offen war schon Nichttreffer).

    Formel: Summe(gewicht * treffer) / Summe(gewicht ueber ALLE Merkmale).
    Rueckgabe: float in [0, 1]. Bei Gewichtssumme 0: 0.0.
    """
    zaehler = 0.0
    nenner = 0.0
    for m in features:
        g = _gewicht(weights, m)
        nenner += g
        if _ist_treffer(auswahl.get(m, OFFEN), typ_profil, m):
            zaehler += g
    return zaehler / nenner if nenner > 0 else 0.0


def abdeckung_anteil(auswahl, features, weights=None):
    """GUETE-KPI 1a: gewichteter Anteil der BEANTWORTETEN Merkmale.
    Beantwortet = nicht OFFEN (eine Ist-Auspraegung liegt vor).
    Formel: Summe(gewicht beantworteter Merkmale) / Summe(gewicht aller).
    Unabhaengig vom Typ. Rueckgabe: float in [0, 1]."""
    beantwortet = 0.0
    gesamt = 0.0
    for m in features:
        g = _gewicht(weights, m)
        gesamt += g
        wert = auswahl.get(m, OFFEN)
        if wert is not OFFEN:
            beantwortet += g
    return beantwortet / gesamt if gesamt > 0 else 0.0


def uebereinstimmung_beantwortet(auswahl, typ_profil, features, weights=None, ausschluss_unit=None):
    """GUETE-KPI 1b / IST-BEREINIGT: Uebereinstimmungsgrad NUR ueber
    'vergleichbare' Merkmale. Vergleichbar (delta*=1) ist ein Merkmal, wenn eine
    Ist-Angabe vorliegt ODER sein Profil komplett ausgeschlossen ist (Blockade).
    OFFENe, nicht blockierte Merkmale bleiben ausgeklammert.
    Ein blockiertes Merkmal zaehlt als Nichttreffer im Nenner -> senkt den Wert.
    Formel: Summe(gewicht * treffer) / Summe(gewicht der vergleichbaren Merkmale).
    Rueckgabe: float in [0, 1]. Wenn nichts vergleichbar: None."""
    zaehler = 0.0
    nenner = 0.0
    for m in features:
        wert = auswahl.get(m, OFFEN)
        blockiert = _ist_blockiert(ausschluss_unit, typ_profil, m)
        # vergleichbar: echte Ist-Angabe ODER blockiert.
        if wert is OFFEN and not blockiert:
            continue  # weder Angabe noch Blockade -> ausklammern
        g = _gewicht(weights, m)
        nenner += g
        # blockierte treffen nie; nur echte Treffer zaehlen im Zaehler.
        if _ist_treffer(wert, typ_profil, m):
            zaehler += g
    if nenner == 0:
        return None
    return zaehler / nenner


def ranking(auswahl, types, features, weights=None):
    """Erstellt die Rangfolge der Typen fuer EINE Einheit nach Haupt-KPI.
    Rueckgabe: Liste von (typ_name, kpi), absteigend sortiert."""
    paare = []
    for typ_name, profil in types.items():
        paare.append((typ_name, uebereinstimmung(auswahl, profil, features, weights)))
    paare.sort(key=lambda x: x[1], reverse=True)
    return paare


def eindeutigkeit(rang):
    """GUETE-KPI 2: Abstand zwischen bestem und zweitbestem Typ.
    rang: Ergebnis von ranking() (absteigend sortiert).
    Rueckgabe: float >= 0. Bei nur einem Typ: None (kein Abstand definierbar)."""
    if len(rang) < 2:
        return None
    return rang[0][1] - rang[1][1]


# --- Soll-Berechnung und harte Verstoesse (Phase 3-Auswertung) ---
# Soll-Menge eines Merkmals = Ist-Auspraegung VEREINIGT mit Potential.
# Ein Soll-Treffer liegt vor, wenn diese Menge das Typprofil schneidet.
# Dadurch ist das Soll nie schlechter als das Ist (alles Ist bleibt enthalten).

def _soll_menge(auswahl, potential_unit, merkmal):
    """Ziel-Menge eines Merkmals: Ist-Auspraegung (falls echte Auspraegung)
    vereinigt mit den Potential-Auspraegungen.
    OFFEN traegt nichts bei (dann zaehlt nur das Potential)."""
    menge = set(potential_unit.get(merkmal, set()))
    ist = auswahl.get(merkmal, OFFEN)
    if ist is not OFFEN:
        menge.add(ist)
    return menge


def uebereinstimmung_soll(auswahl, potential_unit, typ_profil, features, weights=None, ausschluss_unit=None):
    """SOLL-KPI (SOLL-MINIMUM): gewichteter Ziel-Uebereinstimmungsgrad mit EINEM Typ.

    Treffer pro Merkmal, wenn (Ist ∪ Potential) ∩ Typprofil != leer.
    ausschluss_unit optional: blockierte Merkmale sind ohnehin Nichttreffer,
    fuer das Minimum aendert sich der Wert dadurch faktisch nicht.
    Formel: Summe(gewicht * treffer) / Summe(gewicht ueber ALLE Merkmale).
    Rueckgabe: float in [0, 1]. Immer >= Ist-Minimum."""
    zaehler = 0.0
    nenner = 0.0
    for m in features:
        g = _gewicht(weights, m)
        nenner += g
        soll = _soll_menge(auswahl, potential_unit, m)
        if soll & typ_profil.get(m, set()):   # Schnittmenge nicht leer
            zaehler += g
    return zaehler / nenner if nenner > 0 else 0.0


def ranking_soll(auswahl, potential_unit, types, features, weights=None):
    """Rangfolge der Typen nach SOLL-KPI. Liste (typ_name, soll_kpi), absteigend."""
    paare = [(name, uebereinstimmung_soll(auswahl, potential_unit, profil, features, weights))
             for name, profil in types.items()]
    paare.sort(key=lambda x: x[1], reverse=True)
    return paare


# === Dreiteiliges Aehnlichkeitsmass nach GOWER (min / max / bereinigt) =========
# Begriff "angegeben" (delta_i = 1):
#   - Ist:  es liegt eine echte Ist-Auspraegung vor (nicht OFFEN)
#   - Soll: die Vergleichsmenge (Ist ∪ Potenzial) ist nicht leer
# Schon vorhandene Funktionen decken zwei der sechs Werte ab:
#   - Ist-Minimum     = uebereinstimmung()            (OFFEN zaehlt 0, voller Nenner)
#   - Ist-bereinigt   = uebereinstimmung_beantwortet() (nur angegebene Merkmale)
#   - Soll-Minimum    = uebereinstimmung_soll()
# Hier ergaenzt: Ist-Maximum, Soll-Maximum, Soll-bereinigt.

def _ist_angegeben(auswahl, merkmal):
    """delta_i fuer den Ist-Zustand: True, wenn eine echte Ist-Auspraegung vorliegt."""
    wert = auswahl.get(merkmal, OFFEN)
    return wert is not OFFEN


def uebereinstimmung_max(auswahl, typ_profil, features, weights=None, ausschluss_unit=None):
    """IST-MAXIMUM (Best-Case): Merkmale OHNE Angabe zaehlen als Treffer (1).
    ABER: ein Merkmal mit komplett ausgeschlossenem Profil (Blockade-Indikator)
    gilt als 'angegeben' und bekommt KEINEN Best-Case-Bonus, weil es nie ein
    Treffer werden kann. So verspricht das Maximum keine unmoegliche Passung.
    Formel: (Sum gewicht*treffer ueber angegebene + gewicht der nicht angegebenen
             und nicht blockierten) / Sum aller Gewichte."""
    zaehler = 0.0
    nenner = 0.0
    for m in features:
        g = _gewicht(weights, m)
        nenner += g
        blockiert = _ist_blockiert(ausschluss_unit, typ_profil, m)
        # 'angegeben' im erweiterten Sinn: echte Angabe ODER blockiert.
        angegeben_erweitert = _ist_angegeben(auswahl, m) or blockiert
        if not angegeben_erweitert:
            zaehler += g  # nicht angegeben und nicht blockiert -> Best-Case-Treffer
        elif _ist_treffer(auswahl.get(m, OFFEN), typ_profil, m):
            zaehler += g  # echter Treffer (blockierte treffen nie -> zaehlen 0)
    return zaehler / nenner if nenner > 0 else 0.0


def _soll_angegeben(auswahl, potential_unit, merkmal):
    """delta_i fuer den Ziel-Zustand: True, wenn (Ist ∪ Potenzial) nicht leer ist."""
    return len(_soll_menge(auswahl, potential_unit, merkmal)) > 0


def _soll_treffer(auswahl, potential_unit, typ_profil, merkmal):
    """Lokale Ziel-Aehnlichkeit: 1, wenn (Ist ∪ Potenzial) das Profil schneidet."""
    soll = _soll_menge(auswahl, potential_unit, merkmal)
    return len(soll & typ_profil.get(merkmal, set())) > 0


def uebereinstimmung_soll_max(auswahl, potential_unit, typ_profil, features, weights=None, ausschluss_unit=None):
    """SOLL-MAXIMUM (Best-Case): Merkmale ohne Angabe (kein Ist, kein Potenzial)
    zaehlen als Treffer (1). ABER: ein blockiertes Merkmal (Profil komplett
    ausgeschlossen) gilt als 'angegeben' und bekommt KEINEN Best-Case-Bonus."""
    zaehler = 0.0
    nenner = 0.0
    for m in features:
        g = _gewicht(weights, m)
        nenner += g
        blockiert = _ist_blockiert(ausschluss_unit, typ_profil, m)
        angegeben_erweitert = _soll_angegeben(auswahl, potential_unit, m) or blockiert
        if not angegeben_erweitert:
            zaehler += g  # nicht angegeben und nicht blockiert -> Best-Case-Treffer
        elif _soll_treffer(auswahl, potential_unit, typ_profil, m):
            zaehler += g  # echter Soll-Treffer (blockierte treffen nie)
    return zaehler / nenner if nenner > 0 else 0.0


def uebereinstimmung_soll_bereinigt(auswahl, potential_unit, typ_profil, features, weights=None, ausschluss_unit=None):
    """SOLL-BEREINIGT: nur ueber 'vergleichbare' Merkmale. Vergleichbar ist ein
    Merkmal, wenn (Ist ∪ Potenzial) nicht leer ist ODER sein Profil komplett
    ausgeschlossen ist (Blockade). Ein blockiertes Merkmal zaehlt als
    Nichttreffer im Nenner -> senkt den Wert.
    Rueckgabe: float in [0,1] oder None, wenn nichts vergleichbar."""
    zaehler = 0.0
    nenner = 0.0
    for m in features:
        blockiert = _ist_blockiert(ausschluss_unit, typ_profil, m)
        if not _soll_angegeben(auswahl, potential_unit, m) and not blockiert:
            continue
        g = _gewicht(weights, m)
        nenner += g
        if _soll_treffer(auswahl, potential_unit, typ_profil, m):
            zaehler += g
    return zaehler / nenner if nenner > 0 else None


def harte_verstoesse(ausschluss_unit, typ_profil, features):
    """GUETE-KPI 3: Anzahl Merkmale, die fuer DIESEN Typ strukturell
    unerreichbar sind, weil ALLE zulaessigen Auspraegungen ausgeschlossen wurden.

    Bedingung pro Merkmal: das Typprofil ist eine nicht-leere Teilmenge der
    Ausschlussmenge (jede zulaessige Auspraegung ist ausgeschlossen).
    Da eine Ist-Auspraegung nie ausgeschlossen sein kann, ist ein Merkmal mit
    passendem Ist automatisch nicht betroffen.
    Rueckgabe: int >= 0."""
    anzahl = 0
    for m in features:
        profil = typ_profil.get(m, set())
        if not profil:
            continue  # leeres Profil kann nicht verletzt werden
        ausgeschlossen = ausschluss_unit.get(m, set())
        if profil <= ausgeschlossen:   # profil ist Teilmenge von ausgeschlossen
            anzahl += 1
    return anzahl


# --- Zieltyp je Einheit (Schritt 5: Typfestlegung) ---
# --- Engere Auswahl: mehrere Typen je Einheit fuer den Vergleich (Schritt 5) ---
# Bewusst OHNE feste Obergrenze: wie viele Typen sinnvoll sind, haengt von der
# Eindeutigkeit der Bewertung und den verfuegbaren Ressourcen ab. Der Aufwand
# waechst mit jedem weiteren Typ, die Begrenzung liegt daher beim Anwender.

def get_engere_auswahl(state, uid):
    """Liste der Typen in Auswahl-Reihenfolge, die die Einheit zum Vergleich in
    die Zieltypbestimmung nimmt. Leer, falls keine gewaehlt."""
    return state["engere_auswahl"].get(uid, [])


def is_in_engere_auswahl(state, uid, typ_name):
    """True, wenn der Typ fuer die Einheit in der engeren Auswahl ist."""
    return typ_name in state["engere_auswahl"].get(uid, [])


def add_to_engere_auswahl(state, uid, typ_name):
    """Nimmt einen Typ in die engere Auswahl auf und legt seine (leere) eigene
    Ziel-Konfiguration an. Gibt True bei Erfolg zurueck, False wenn die Auswahl
    den Typ bereits enthaelt."""
    liste = state["engere_auswahl"].setdefault(uid, [])
    if typ_name in liste:
        return False
    liste.append(typ_name)
    state["soll"].setdefault(uid, {}).setdefault(typ_name, {})
    return True


def remove_from_engere_auswahl(state, uid, typ_name):
    """Entfernt einen Typ aus der engeren Auswahl samt seiner Ziel-Konfiguration
    und seiner Aufwaende. War er der finale Zieltyp, wird auch dieser zurueckgesetzt."""
    liste = state["engere_auswahl"].get(uid, [])
    if typ_name in liste:
        liste.remove(typ_name)
    state["soll"].get(uid, {}).pop(typ_name, None)
    state["aufwand"].get(uid, {}).pop(typ_name, None)
    state["kosten"].get(uid, {}).pop(typ_name, None)
    state["dauer"].get(uid, {}).pop(typ_name, None)
    if state["zieltyp"].get(uid) == typ_name:
        state["zieltyp"][uid] = None


# --- Finaler Zieltyp je Einheit (nach dem Vergleich festgelegt) ---
def set_zieltyp(state, uid, typ_name):
    """Legt den FINALEN Zieltyp einer Einheit fest. typ_name=None loest die
    Festlegung. Ein finaler Zieltyp gehoert stets zur engeren Auswahl; ist er
    noch nicht enthalten, wird er (sofern Platz) aufgenommen."""
    if typ_name is not None:
        add_to_engere_auswahl(state, uid, typ_name)  # idempotent, falls schon drin
    state["zieltyp"][uid] = typ_name


def get_zieltyp(state, uid):
    """Liefert den finalen Zieltyp oder None (noch nicht festgelegt)."""
    return state["zieltyp"].get(uid)


# --- Soll-Auspraegung je Merkmal UND Typ (Schritt 6: Detaillierung) ---
def set_soll(state, uid, typ_name, merkmal, auspraegung):
    """Setzt die angestrebte Ziel-Auspraegung eines Merkmals fuer EINEN Typ
    (oder OFFEN). Jeder Typ der engeren Auswahl hat eine eigene Ziel-Konfiguration.
    Aendert sich das Ziel, wird der zugehoerige Aufwand verworfen, da er sich auf
    die zuvor gewaehlte Auspraegung bezog."""
    soll_typ = state["soll"].setdefault(uid, {}).setdefault(typ_name, {})
    if soll_typ.get(merkmal, OFFEN) != auspraegung:
        state["aufwand"].get(uid, {}).get(typ_name, {}).pop(merkmal, None)
        state["kosten"].get(uid, {}).get(typ_name, {}).pop(merkmal, None)
        state["dauer"].get(uid, {}).get(typ_name, {}).pop(merkmal, None)
    soll_typ[merkmal] = auspraegung


def get_soll(state, uid, typ_name, merkmal):
    """Liefert die angestrebte Ziel-Auspraegung eines Merkmals fuer EINEN Typ,
    oder OFFEN (nicht festgelegt)."""
    return state["soll"].get(uid, {}).get(typ_name, {}).get(merkmal, OFFEN)


# --- Aufwand je Merkmal und Typ (Schritt 6: Detaillierung) ---
AUFWAND_GERING = 1
AUFWAND_MITTEL = 2
AUFWAND_HOCH = 3
AUFWAND_STUFEN = (AUFWAND_GERING, AUFWAND_MITTEL, AUFWAND_HOCH)
# Stufe 0 ist keine erfassbare Aufwandsangabe, sondern die Betrachtungsstufe
# "ohne jede Aenderung". Sie wird nur in der gestaffelten Auswertung gebraucht.
AUFWAND_KEIN = 0
AUFWAND_STUFEN_K = (AUFWAND_KEIN, AUFWAND_GERING, AUFWAND_MITTEL, AUFWAND_HOCH)
AUFWAND_LABEL = {AUFWAND_GERING: "gering", AUFWAND_MITTEL: "mittel", AUFWAND_HOCH: "hoch"}

# Phasen der Maßnahmenplanung. Sie werden anfangs aus dem Aufwand vorgeschlagen,
# sind danach aber frei gebildet und daher als eigene Groesse gefuehrt. Die Werte
# stimmen zahlenmaessig mit den Aufwandsstufen ueberein, tragen aber keine
# Aufwandsbedeutung mehr.
PHASEN = (1, 2, 3)


def set_aufwand(state, uid, typ_name, merkmal, wert):
    """Setzt den geschaetzten Aufwand (1=gering, 2=mittel, 3=hoch), um das Merkmal
    vom heutigen Zustand in die gewaehlte Ziel-Auspraegung zu bringen, fuer EINEN Typ."""
    state["aufwand"].setdefault(uid, {}).setdefault(typ_name, {})[merkmal] = wert


def get_aufwand(state, uid, typ_name, merkmal):
    """Liefert den geschaetzten Aufwand oder None (noch nicht geschaetzt)."""
    return state["aufwand"].get(uid, {}).get(typ_name, {}).get(merkmal)


def set_kosten(state, uid, typ_name, merkmal, wert):
    """Optionale Kostenschaetzung (Euro, float) oder None fuer ein Merkmal
    zu einem Typ. None loescht die Angabe."""
    if wert is None:
        state["kosten"].get(uid, {}).get(typ_name, {}).pop(merkmal, None)
        return
    state["kosten"].setdefault(uid, {}).setdefault(typ_name, {})[merkmal] = float(wert)


def get_kosten(state, uid, typ_name, merkmal):
    """Kostenschaetzung (Euro) oder None, wenn nicht erfasst."""
    return state.get("kosten", {}).get(uid, {}).get(typ_name, {}).get(merkmal)


def set_dauer(state, uid, typ_name, merkmal, wert):
    """Optionale Dauerschaetzung (Wochen, float) oder None fuer ein Merkmal zu
    einem Typ. None loescht die Angabe."""
    if wert is None:
        state["dauer"].get(uid, {}).get(typ_name, {}).pop(merkmal, None)
        return
    state["dauer"].setdefault(uid, {}).setdefault(typ_name, {})[merkmal] = float(wert)


def get_dauer(state, uid, typ_name, merkmal):
    """Dauerschaetzung (Wochen) oder None, wenn nicht erfasst."""
    return state.get("dauer", {}).get(uid, {}).get(typ_name, {}).get(merkmal)


def soll_kandidaten(status, optionen):
    """Auspraegungen mit Wahl-Knopf in der Zieltypbestimmung. In Fall 1 (das Ist liegt
    bereits im Profil) gibt es keine Wahl, da das Merkmal erfuellt ist und die
    Ziel-Ausprägung automatisch dem Ist entspricht. In allen uebrigen Faellen (2 bis
    5) stehen alle Auspraegungen zur freien Wahl. Damit kann der Anwender bewusst
    auch abweichen, etwa um im Sinne einer betrachtungseinheitsuebergreifenden
    Synergie dieselbe Ausprägung wie eine andere Einheit anzustreben. Eine
    profilfremde Wahl bildet dann kein Handlungsfeld, sondern erscheint als nicht
    erreichtes Merkmal. Rueckgabe als Menge."""
    if status == STATUS_IST:
        return set()
    return set(optionen)


def soll_vorbelegen(state, uid, typ_name, features):
    """Belegt das Ziel mit dem Ist vor, wo es ein Ist gibt (Faelle 1, 2, 3) und
    noch nichts gewaehlt wurde. Faelle ohne Ist (4, 5) bleiben offen und muessen
    aktiv entschieden werden. Idempotent; sorgt dafuer, dass auch nicht geoeffnete
    Tabs vollstaendig vorbelegt sind."""
    for m in features:
        if get_soll(state, uid, typ_name, m) is not OFFEN:
            continue
        ist = get_choice(state, uid, m)
        if ist is not OFFEN:
            set_soll(state, uid, typ_name, m, ist)


def detaillierung_vollstaendig(state, uid, typ_name, features):
    """True, wenn fuer JEDES Merkmal eine Entscheidung vorliegt: ein Ziel (eine
    Auspraegung oder bewusst NICHTS_ANSTREBEN) und, falls das Ziel vom Ist abweicht,
    ein erfasster Aufwand. Setzt voraus, dass zuvor soll_vorbelegen gelaufen ist."""
    for m in features:
        soll = get_soll(state, uid, typ_name, m)
        if soll is OFFEN:
            return False
        ist = get_choice(state, uid, m)
        if soll != NICHTS_ANSTREBEN and soll != ist:
            if get_aufwand(state, uid, typ_name, m) is None:
                return False
    return True


def alle_detaillierungen_vollstaendig(state, features):
    """True, wenn mindestens eine Einheit eine engere Auswahl hat und alle Typen
    aller engeren Auswahlen vollstaendig detailliert sind. Vor der Pruefung wird
    das Ziel ueberall vorbelegt, damit nicht geoeffnete Tabs nicht blockieren."""
    hat_auswahl = False
    for uid in state["units"]:
        for typ in get_engere_auswahl(state, uid):
            hat_auswahl = True
            soll_vorbelegen(state, uid, typ, features)
            if not detaillierung_vollstaendig(state, uid, typ, features):
                return False
    return hat_auswahl


def soll_score_gestaffelt(state, uid, typ_name, typ_profil, features, weights=None):
    """Ziel-Score je Aufwandsstufe fuer EINEN Typ, als dict {0: .., 1: .., 2: .., 3: ..}.

    Stufe 0 ist die bereits ohne jede Aenderung erreichte Uebereinstimmung.

    Ein Merkmal ist ein Treffer, wenn seine Ziel-Auspraegung im Profil liegt. Der
    Erreichungsaufwand ist 0, wenn das Ziel bereits dem Ist entspricht (schon
    erreicht, kein Aufwand), sonst der erfasste Aufwand (1..3). Ein Treffer zaehlt
    in Stufe k, wenn sein Erreichungsaufwand <= k. Der Score bei Stufe k ist damit
    die erreichte gewichtete Profilkonformitaet, wenn alle Aenderungen bis
    einschliesslich Aufwand k umgesetzt werden. Stufe 3 ist der volle Ziel-Score.

    Nicht getroffene Merkmale (Ist belassen und nicht profilkonform, bewusst
    nichts angestrebt, oder eine nicht profilkonforme Auspraegung angestrebt)
    zaehlen in keiner Stufe."""
    gesamt = sum(_gewicht(weights, m) for m in features)
    if gesamt == 0:
        return {k: 0.0 for k in AUFWAND_STUFEN_K}
    treffer = {k: 0.0 for k in AUFWAND_STUFEN_K}
    for m in features:
        soll = get_soll(state, uid, typ_name, m)
        profil = typ_profil.get(m, set())
        if soll is OFFEN or soll == NICHTS_ANSTREBEN or soll not in profil:
            continue  # kein Treffer
        ist = get_choice(state, uid, m)
        if soll == ist:
            stufe = 0  # schon erreicht
        else:
            stufe = get_aufwand(state, uid, typ_name, m)
            if stufe is None:
                stufe = AUFWAND_HOCH  # unvollstaendig: konservativ erst bei voller Umsetzung
        w = _gewicht(weights, m)
        for k in AUFWAND_STUFEN_K:
            if stufe <= k:
                treffer[k] += w
    return {k: treffer[k] / gesamt for k in AUFWAND_STUFEN_K}


def aufwand_verteilung(state, uid, typ_name, typ_profil, features):
    """Zaehlt, wie viele Merkmale je Aufwandsstufe tatsaechlich zu einer
    profilkonformen Annaeherung fuehren (Treffer mit Ziel != Ist). Rueckgabe:
    dict {1: n, 2: n, 3: n}. Nuetzlich, um neben dem Score zu zeigen, wie viele
    Aenderungen jede Stufe kostet."""
    zahl = {AUFWAND_GERING: 0, AUFWAND_MITTEL: 0, AUFWAND_HOCH: 0}
    for m in features:
        soll = get_soll(state, uid, typ_name, m)
        profil = typ_profil.get(m, set())
        if soll is OFFEN or soll == NICHTS_ANSTREBEN or soll not in profil:
            continue
        ist = get_choice(state, uid, m)
        if soll == ist:
            continue  # schon erreicht, kostet nichts
        stufe = get_aufwand(state, uid, typ_name, m)
        if stufe in zahl:
            zahl[stufe] += 1
    return zahl


# --- Merkmals-Status je Typ (fuer die Statusanzeige im Dashboard) ---
# Fuenf Stati beschreiben, wie ein Merkmal der Einheit zu EINEM Typ steht.
# Sie entsprechen genau den fuenf Faellen aus Kapitel 5.3.2:
STATUS_IST = "ist"                     # Fall 1: Ist-Auspraegung liegt im Profil
STATUS_POTENTIAL = "potential"         # Fall 2: nicht im Ist, aber ueber Potential erreichbar
STATUS_IST_UNPASSEND = "ist_unpassend" # Fall 3: Ist vorhanden, trifft nicht, kein Potential trifft
STATUS_OFFEN = "offen"                 # Fall 4: kein Ist, nicht blockiert (echt offen)
STATUS_BLOCKIERT = "blockiert"         # Fall 5: Profil vollstaendig ausgeschlossen

def merkmal_status(auswahl, potential_unit, ausschluss_unit, typ_profil, merkmal):
    """Bestimmt den Vergleichsfall EINES Merkmals der Einheit gegenueber EINEM Typ.
    Rueckgabe: einer der STATUS_*-Werte.

    Logik (Reihenfolge wichtig):
      1) Ist-Auspraegung im Profil                     -> STATUS_IST
      2) (Ist ∪ Potential) schneidet Profil            -> STATUS_POTENTIAL
      3) alle zulaessigen Auspraegungen ausgeschlossen -> STATUS_BLOCKIERT
      4) Ist vorhanden, trifft aber nicht              -> STATUS_IST_UNPASSEND
      5) sonst (kein Ist, nicht blockiert)             -> STATUS_OFFEN
    """
    profil = typ_profil.get(merkmal, set())
    ist = auswahl.get(merkmal, OFFEN)

    # 1) Ist trifft direkt?
    if ist is not OFFEN and ist in profil:
        return STATUS_IST

    # 2) Ueber Potential (inkl. Ist) erreichbar?
    soll = _soll_menge(auswahl, potential_unit, merkmal)
    if soll & profil:
        return STATUS_POTENTIAL

    # 3) Strukturell blockiert? (Profil nicht leer und ganz ausgeschlossen)
    ausgeschlossen = ausschluss_unit.get(merkmal, set())
    if profil and profil <= ausgeschlossen:
        return STATUS_BLOCKIERT

    # 4) Ist vorhanden, trifft aber nicht -> Fall 3 (nichts trifft)
    if ist is not OFFEN:
        return STATUS_IST_UNPASSEND

    # 5) Kein Ist, nicht blockiert -> echt offenes Merkmal (Fall 4)
    return STATUS_OFFEN


# ====================================================== Schritt 4: Massnahmenplanung
# Die Handlungsfelder ergeben sich unmittelbar aus dem Vergleich von Ist- und
# Soll-Ausprägung des festgelegten Zieltyps. Der Anwender ergaenzt lediglich die
# konkrete Massnahme, die Phase und die Verantwortlichkeit.

def handlungsfelder(state, uid, typ_name, typ_profil, features, weights=None):
    """Liste der Handlungsfelder eines Zieltyps. Ein Handlungsfeld erfuellt ZWEI
    Bedingungen: die Ziel-Auspraegung weicht vom Ist ab (es gibt eine Veraenderung)
    UND sie liegt im Profil des Zieltyps (die Veraenderung fuehrt naeher an den Typ
    heran). Nur dann dient die Veraenderung der Annaeherung an den Zieltyp.

    Kein Handlungsfeld sind daher: Merkmale mit Ziel = Ist (keine Veraenderung),
    Merkmale ohne angestrebte Ausprägung (bewusst nichts) und Merkmale mit
    profilfremder Ziel-Ausprägung. Die beiden letzten sind bewusste Abweichungen
    vom Zieltyp und werden ueber nicht_erreichte_merkmale gesondert ausgewiesen.

    Rueckgabe: Liste von dicts mit merkmal, ist, soll, aufwand, gewicht, gewinn
    und ist_bekannt. Der Gewinn ist der Ähnlichkeitsgewinn, also der Anteil, um den
    die globale Ähnlichkeit steigt, wenn das Handlungsfeld erreicht wird. Da ein
    Handlungsfeld profilkonform ist und im Ist nicht erfuellt war, entspricht der
    lokale Zuwachs stets eins, sodass der Gewinn dem normierten Gewicht des Merkmals
    gleicht (w_i geteilt durch die Summe aller Merkmalsgewichte). Sortiert nach
    Aufwand, dann nach absteigendem Gewinn, sodass wirksame und schnell umsetzbare
    Punkte oben stehen."""
    gesamtgewicht = sum(_gewicht(weights, m) for m in features)
    felder = []
    for m in features:
        soll = get_soll(state, uid, typ_name, m)
        if soll is OFFEN or soll == NICHTS_ANSTREBEN:
            continue                      # keine Ausprägung angestrebt
        if soll not in typ_profil.get(m, set()):
            continue                      # profilfremdes Soll, bewusste Abweichung
        ist = get_choice(state, uid, m)
        if soll == ist:
            continue                      # bereits erreicht, kein Handlungsfeld
        gewicht = _gewicht(weights, m)
        felder.append({
            "merkmal": m,
            "ist": ist,                   # OFFEN, wenn Ist unbekannt war
            "soll": soll,
            "aufwand": get_aufwand(state, uid, typ_name, m),
            "gewicht": gewicht,
            "gewinn": gewicht / gesamtgewicht if gesamtgewicht else 0.0,
            "ist_bekannt": ist is not OFFEN,
        })
    felder.sort(key=lambda f: ((f["aufwand"] or AUFWAND_HOCH), -f["gewinn"]))
    return felder


EXAKT_GRENZE = 18   # bis so viele bezifferte Handlungsfelder je Typ exakt rechnen


def _pareto_filter(punkte):
    """Behaelt aus (kosten, gewinn, merkmale)-Punkten nur die nicht dominierten:
    aufsteigend nach Kosten sortiert, streng steigender Gewinn."""
    punkte.sort(key=lambda p: (p[0], -p[1]))
    front, best = [], -1.0
    for k, g, ms in punkte:
        if g > best + 1e-12:
            front.append((k, g, ms))
            best = g
    return front


def budgetkurve(state, uid, typ_name, typ_profil, features, weights=None):
    """Datenpunkte der Budgetkurve fuer EINEN Typ: die exakt hoechste erreichbare
    Ziel-Aehnlichkeit ueber dem Budget. Grundlage sind die Handlungsfelder mit
    erfasster Kostenschaetzung. Berechnet wird die Effizienzgrenze (Pareto-Front)
    ueber alle Kombinationen, sodass zu jedem Budget die tatsaechlich beste Auswahl
    gezeigt wird. Bis EXAKT_GRENZE bezifferte Handlungsfelder geschieht das exakt per
    dynamischer Programmierung; darueber (oder wenn die Front zu gross wird) greift
    die gierige Naeherung nach Gewinn je Euro, damit die Rechenzeit beschraenkt
    bleibt. Handlungsfelder ohne Kosten fliessen nicht in die Kurve ein und deckeln
    sie.

    Rueckgabe (dict):
      basis         Aehnlichkeit ohne jede Aenderung (Budget 0),
      punkte        Liste (budget, aehnlichkeit, merkmale) der Effizienzgrenze;
                    merkmale ist das Tupel der in dieser Auswahl enthaltenen
                    Merkmale (leer beim Basispunkt),
      deckel        Aehnlichkeit, wenn alle bezifferten Handlungsfelder umgesetzt
                    sind,
      unbeziffert   Liste der Handlungsfelder ohne Kosten,
      gewinn_offen  Summe der Gewinne der unbezifferten Handlungsfelder,
      max_gesamt    deckel + gewinn_offen,
      exakt         True, wenn die Kurve exakt gerechnet wurde, sonst False.
    """
    felder = handlungsfelder(state, uid, typ_name, typ_profil, features, weights)
    score = soll_score_gestaffelt(state, uid, typ_name, typ_profil, features, weights)
    basis = score[AUFWAND_KEIN]

    mit_kosten, ohne_kosten = [], []
    for f in felder:
        k = get_kosten(state, uid, typ_name, f["merkmal"])
        if k is None:
            ohne_kosten.append(f)
        else:
            mit_kosten.append((k, f))

    def _exakte_front():
        # Pareto-Front per DP: jede (kosten, gewinn, merkmale)-Kombination, nach
        # jedem Handlungsfeld auf die nicht dominierten Punkte reduziert.
        front = [(0.0, 0.0, ())]
        for k, f in mit_kosten:
            erweitert = []
            for kk, gg, ms in front:
                erweitert.append((kk, gg, ms))
                erweitert.append((kk + k, gg + f["gewinn"], ms + (f["merkmal"],)))
            front = _pareto_filter(erweitert)
            if len(front) > 200000:      # Sicherheitsgrenze -> Naeherung
                return None
        front.sort()
        return front

    front = _exakte_front() if len(mit_kosten) <= EXAKT_GRENZE else None
    exakt = front is not None
    if exakt:
        punkte = [(kk, basis + gg, ms) for kk, gg, ms in front]
        deckel = basis + sum(f["gewinn"] for _, f in mit_kosten)
    else:
        mit_kosten.sort(
            key=lambda kf: (kf[1]["gewinn"] / kf[0]) if kf[0] > 0 else -1.0,
            reverse=True)
        punkte = [(0.0, basis, ())]
        budget, aehn, ms = 0.0, basis, ()
        for k, f in mit_kosten:
            budget += k
            aehn += f["gewinn"]
            ms = ms + (f["merkmal"],)
            punkte.append((budget, aehn, ms))
        deckel = aehn

    gewinn_offen = sum(f["gewinn"] for f in ohne_kosten)
    return {
        "basis": basis,
        "punkte": punkte,
        "deckel": deckel,
        "unbeziffert": ohne_kosten,
        "gewinn_offen": gewinn_offen,
        "max_gesamt": deckel + gewinn_offen,
        "exakt": exakt,
    }
def budget_optimum(state, uid, typ_name, typ_profil, features, weights, budget):
    """Exakte Loesung des Auswahlproblems (Rucksack) fuer EINEN Typ: waehlt aus den
    Merkmalen mit erfasster Kostenschaetzung die Kombination, die das Budget nicht
    ueberschreitet und die hoechste Ziel-Aehnlichkeit erreicht. Merkmale ohne
    Kosten bleiben aussen vor.

    Anders als die (gierige) Budgetkurve ist dies das echte Optimum fuer den einen
    Budgetwert; beide koennen daher in seltenen Faellen leicht abweichen.

    Rueckgabe (dict): aehnlichkeit (erreichbar), merkmale (Liste der gewaehlten),
    kosten (Summe). Bis zu 16 bezifferte Merkmale werden vollstaendig enumeriert,
    darueber greift die gierige Naeherung.
    """
    felder = handlungsfelder(state, uid, typ_name, typ_profil, features, weights)
    mit_kosten = [(get_kosten(state, uid, typ_name, f["merkmal"]), f) for f in felder]
    mit_kosten = [(k, f) for k, f in mit_kosten if k is not None]
    basis = soll_score_gestaffelt(state, uid, typ_name, typ_profil,
                                  features, weights)[AUFWAND_KEIN]
    n = len(mit_kosten)

    if n <= 16:
        beste_gewinn, beste_maske = 0.0, 0
        for maske in range(1 << n):
            ksum = gsum = 0.0
            for i in range(n):
                if maske & (1 << i):
                    ksum += mit_kosten[i][0]
                    gsum += mit_kosten[i][1]["gewinn"]
            if ksum <= budget and gsum > beste_gewinn:
                beste_gewinn, beste_maske = gsum, maske
        gewaehlt = [mit_kosten[i][1] for i in range(n) if beste_maske & (1 << i)]
        kosten = sum(mit_kosten[i][0] for i in range(n) if beste_maske & (1 << i))
    else:
        rang = sorted(mit_kosten, key=lambda kf: kf[1]["gewinn"] / kf[0],
                      reverse=True)
        gewaehlt, kosten, rest = [], 0.0, budget
        for k, f in rang:
            if k <= rest:
                gewaehlt.append(f)
                kosten += k
                rest -= k

    return {
        "aehnlichkeit": basis + sum(f["gewinn"] for f in gewaehlt),
        "merkmale": [f["merkmal"] for f in gewaehlt],
        "kosten": kosten,
    }


def nicht_erreichte_merkmale(state, uid, typ_name, typ_profil, features):
    """Merkmale, in denen der Zieltyp dauerhaft nicht erreicht wird, weil das Ziel
    nicht im Profil liegt oder bewusst nichts angestrebt wurde. Sie begruenden,
    warum die erreichbare Ähnlichkeit unter eins bleibt."""
    offen = []
    for m in features:
        soll = get_soll(state, uid, typ_name, m)
        if soll == NICHTS_ANSTREBEN:
            offen.append((m, "bewusst nichts angestrebt"))
        elif soll is OFFEN:
            offen.append((m, "keine Ziel-Angabe"))
        elif soll not in typ_profil.get(m, set()):
            offen.append((m, "Ziel liegt ausserhalb des Profils"))
    return offen


def synergie_potenzial(state, uid, merkmal, soll):
    """Sucht andere Betrachtungseinheiten, die im selben Merkmal dieselbe
    Ziel-Ausprägung wie das betrachtete Handlungsfeld verfolgen. Verglichen werden
    die Ziel-Ausprägungen aller Einheiten, nicht nur ihre Handlungsfelder, damit
    auch eine Einheit erfasst wird, welche die Ausprägung bereits besitzt.

    Rueckgabe: Liste von (uid2, besitzt_bereits), aufsteigend nach uid2.
    besitzt_bereits ist True, wenn die andere Einheit die Ausprägung schon im Ist
    hat und damit als Vorbild dienen kann, sonst False (sie strebt sie ebenfalls an).
    Der Parameter soll ist die Ziel-Ausprägung des Handlungsfeldes und stets eine
    echte Profilausprägung."""
    treffer = []
    for uid2 in state["units"]:
        if uid2 == uid:
            continue
        typ2 = get_zieltyp(state, uid2)
        if not typ2:
            continue                       # ohne Zieltyp keine Soll-Konfiguration
        if get_soll(state, uid2, typ2, merkmal) == soll:
            besitzt = get_choice(state, uid2, merkmal) == soll
            treffer.append((uid2, besitzt))
    treffer.sort(key=lambda t: t[0])
    return treffer


# ---------------------------------------- Unternehmensweite Zusammenfuehrung
def alle_handlungsfelder(state, types, features, weights=None):
    """Sammelt die Handlungsfelder aller Betrachtungseinheiten mit festgelegtem
    Zieltyp. Jedes Feld wird um den Schluessel uid ergaenzt. Sortiert nach der
    Merkmalsreihenfolge, dann nach Einheit."""
    reihenfolge = {m: i for i, m in enumerate(features)}
    felder = []
    for uid in state["units"]:
        typ = get_zieltyp(state, uid)
        if not typ:
            continue
        for f in handlungsfelder(state, uid, typ, types[typ], features, weights):
            g = dict(f)
            g["uid"] = uid
            felder.append(g)
    felder.sort(key=lambda f: (reihenfolge.get(f["merkmal"], 10 ** 9), f["uid"]))
    return felder


def buendel_kandidaten(state, types, features, weights=None):
    """Kombinationen aus Merkmal und Ziel-Ausprägung, die bei mehreren Einheiten
    zugleich ein Handlungsfeld sind und sich daher zusammenlegen lassen.
    Rueckgabe: dict {(merkmal, soll): [uid, ...]} mit mindestens zwei Einheiten."""
    from collections import OrderedDict
    gruppen = OrderedDict()
    for f in alle_handlungsfelder(state, types, features, weights):
        gruppen.setdefault((f["merkmal"], f["soll"]), []).append(f["uid"])
    return {k: v for k, v in gruppen.items() if len(v) >= 2}


def ist_gebuendelt(state, merkmal, soll):
    """True, wenn der Anwender diese Kombination zusammengelegt hat."""
    return (merkmal, soll) in state.get("buendel", set())


def toggle_buendel(state, merkmal, soll):
    """Legt eine Kombination aus Merkmal und Ziel-Ausprägung zusammen oder loest die
    Zusammenlegung wieder auf. Das Zusammenlegen ist damit umkehrbar."""
    b = state.setdefault("buendel", set())
    schluessel = (merkmal, soll)
    if schluessel in b:
        b.discard(schluessel)
        state.get("buendel_werte", {}).pop(schluessel, None)
    else:
        b.add(schluessel)


def set_buendel_werte(state, merkmal, soll, aufwand, kosten, dauer):
    """Speichert die gemeinsamen Werte einer zusammengelegten Kombination. Kosten
    und Dauer duerfen None sein, der Aufwand ist die gemeinsame Stufe."""
    state.setdefault("buendel_werte", {})[(merkmal, soll)] = {
        "aufwand": aufwand,
        "kosten": None if kosten is None else float(kosten),
        "dauer": None if dauer is None else float(dauer),
    }


def get_buendel_werte(state, merkmal, soll):
    """Gemeinsame Werte einer Zusammenlegung (dict mit aufwand/kosten/dauer) oder
    None, wenn noch keine erfasst wurden."""
    return state.get("buendel_werte", {}).get((merkmal, soll))


def einzelwerte(state, einheiten, merkmal):
    """Sammelt die erfassten Aufwaende, Kosten und Dauern der beteiligten Einheiten
    fuer ein Merkmal (jeweils bezogen auf den Zieltyp der Einheit)."""
    au, ks, ds = [], [], []
    for uid in einheiten:
        typ = get_zieltyp(state, uid)
        a = get_aufwand(state, uid, typ, merkmal)
        k = get_kosten(state, uid, typ, merkmal)
        d = get_dauer(state, uid, typ, merkmal)
        if a in AUFWAND_STUFEN:
            au.append(a)
        if k is not None:
            ks.append(k)
        if d is not None:
            ds.append(d)
    return au, ks, ds


def zeile_werte(state, zeile):
    """Aufwand, Kosten und Dauer einer unternehmensweiten Zeile. Bei einer
    Zusammenlegung mit erfassten gemeinsamen Werten diese, sonst der Aufwand aus der
    Zeile und der hoechste Kosten- bzw. Dauerwert der beteiligten Einheiten. None,
    wo nichts vorliegt."""
    if zeile.get("gebuendelt"):
        bw = get_buendel_werte(state, zeile["merkmal"], zeile["soll"])
        if bw:
            return bw.get("aufwand"), bw.get("kosten"), bw.get("dauer")
    _, ks, ds = einzelwerte(state, zeile["einheiten"], zeile["merkmal"])
    return zeile["aufwand"], (max(ks) if ks else None), (max(ds) if ds else None)


def synergie_gruppen(state):
    """Liste der inhaltlichen Synergien. Jede Gruppe ist ein dict mit 'felder'
    (frozenset von Handlungsfeld-Schluesseln), 'aufwand', 'kosten' und 'dauer'."""
    return state.setdefault("synergien", [])


def synergie_von(state, schluessel):
    """Die Synergie-Gruppe, die einen Handlungsfeld-Schluessel enthaelt, sonst
    None. Ein Handlungsfeld gehoert zu hoechstens einer Synergie."""
    for g in state.get("synergien", []):
        if schluessel in g["felder"]:
            return g
    return None


def add_synergie(state, schluessel_menge, aufwand, kosten, dauer):
    """Legt eine neue inhaltliche Synergie aus mehreren Handlungsfeld-Schluesseln
    an, mit gemeinsamem Aufwand sowie optionalen gemeinsamen Kosten und Dauer."""
    state.setdefault("synergien", []).append({
        "felder": frozenset(schluessel_menge),
        "aufwand": aufwand,
        "kosten": None if kosten is None else float(kosten),
        "dauer": None if dauer is None else float(dauer),
    })


def remove_synergie(state, index):
    """Loest die Synergie-Gruppe an der gegebenen Position wieder auf."""
    g = state.get("synergien", [])
    if 0 <= index < len(g):
        g.pop(index)


def remove_synergie_by_id(state, mid):
    """Loest die Synergie auf, deren Feldmenge der gegebenen id (frozenset)
    entspricht. Nutzt man, wenn die Synergie ueber ihre Massnahmen-id angesprochen
    wird statt ueber ihre Position."""
    g = state.get("synergien", [])
    for i, gr in enumerate(g):
        if gr["felder"] == mid:
            g.pop(i)
            return


def massnahmen_liste(state, types, features, weights=None):
    """Bildet aus den unternehmensweiten Zeilen und den inhaltlichen Synergien die
    endgueltigen Massnahmen. Eine Zeile ohne Synergie ist eine eigene Massnahme,
    eine Synergie fasst mehrere Zeilen zu einer Massnahme zusammen (Gewinn summiert,
    Aufwand/Kosten/Dauer aus den gemeinsamen Werten der Synergie).

    Jede Massnahme ist ein dict mit: id (hashbar), merkmale (Liste), einheiten
    (Liste von uids), gewinn (Anteil, ggf. summiert), aufwand, kosten, dauer,
    ist_synergie (bool). Reihenfolge: erst Synergien, dann uebrige Zeilen in
    Zeilenreihenfolge.
    """
    zeilen = unternehmensweite_uebersicht(state, types, features, weights)
    per_key = {hf_schluessel(z): z for z in zeilen}

    massnahmen, vergeben = [], set()
    for g in state.get("synergien", []):
        felder = [k for k in g["felder"] if k in per_key]
        if len(felder) < 2:
            continue
        vergeben.update(felder)
        zs = [per_key[k] for k in felder]
        einheiten = []
        for z in zs:
            einheiten.extend(z["einheiten"])
        massnahmen.append({
            "id": frozenset(felder),
            "merkmale": [z["merkmal"] for z in zs],
            "ziele": [z["soll"] for z in zs],
            "ist_texte": ["/".join(sorted(set(str(x) for x in z["ist_werte"])))
                          for z in zs],
            "einheiten": list(dict.fromkeys(einheiten)),
            "gewinn": sum(z["gewinn"] for z in zs),
            "aufwand": g.get("aufwand"),
            "kosten": g.get("kosten"),
            "dauer": g.get("dauer"),
            "ist_synergie": True,
        })
    for k, z in per_key.items():
        if k in vergeben:
            continue
        aufw, kosten, dauer = zeile_werte(state, z)
        massnahmen.append({
            "id": k,
            "merkmale": [z["merkmal"]],
            "ziele": [z["soll"]],
            "ist_texte": ["/".join(sorted(set(str(x) for x in z["ist_werte"])))],
            "einheiten": list(z["einheiten"]),
            "gewinn": z["gewinn"],
            "aufwand": aufw,
            "kosten": kosten,
            "dauer": dauer,
            "ist_synergie": False,
        })
    return massnahmen


def abhaengigkeiten(state):
    """Menge der zwingenden Abhaengigkeiten als (vorher_id, nachher_id): die erste
    Massnahme muss vor der zweiten umgesetzt werden."""
    return state.setdefault("abhaengigkeiten", set())


def set_abhaengigkeit(state, vorher, nachher):
    """Legt fest, dass die Massnahme 'vorher' zwingend vor 'nachher' liegt. Die
    Gegenrichtung wird entfernt, damit ein Paar nicht widerspruechlich wird."""
    a = state.setdefault("abhaengigkeiten", set())
    a.discard((nachher, vorher))
    a.add((vorher, nachher))


def clear_abhaengigkeit(state, a_id, b_id):
    """Entfernt beide Richtungen einer Abhaengigkeit zwischen zwei Massnahmen."""
    a = state.setdefault("abhaengigkeiten", set())
    a.discard((a_id, b_id))
    a.discard((b_id, a_id))


def abhaengigkeit_status(state, a_id, b_id):
    """Status des geordneten Paars (a, b): 'a_vor_b', 'b_vor_a' oder 'keine'."""
    a = state.get("abhaengigkeiten", set())
    if (a_id, b_id) in a:
        return "a_vor_b"
    if (b_id, a_id) in a:
        return "b_vor_a"
    return "keine"


def vorgaenger(state, mid, gueltige_ids):
    """Menge der Massnahmen-ids, die zwingend vor 'mid' liegen muessen und in der
    Menge der aktuell gueltigen ids enthalten sind (nicht mehr existierende
    Massnahmen werden ignoriert)."""
    return {v for (v, n) in state.get("abhaengigkeiten", set())
            if n == mid and v in gueltige_ids}


def einheit_kennzahlen(state, uid, types, features, weights=None):
    """Aufwand, Kosten, Dauer und Aehnlichkeitsgewinn je Einheit als Grundlage der
    Einheiten-Priorisierung. Aufwand, Kosten und Dauer werden aus den Massnahmen
    abgeleitet, damit Zusammenlegungen wirken: eine gemeinsame Massnahme wird den
    beteiligten Einheiten zu gleichen Teilen zugerechnet (Wert geteilt durch die
    Zahl der Einheiten). Fehlende Kosten oder Dauern werden nicht als 0 gewertet,
    sondern aus dem Durchschnitt der bezifferten Massnahmen der Einheit auf deren
    Gesamtzahl hochgerechnet, damit gut dokumentierte Einheiten nicht benachteiligt
    werden. Der Ähnlichkeitsgewinn bleibt einheitenspezifisch (Summe der
    Handlungsfelder), weil er von der Zusammenlegung unberuehrt bleibt. Zusaetzlich
    wird 'kosten_vollstaendig' als Anteil bezifferter Massnahmen zurueckgegeben."""
    typ = get_zieltyp(state, uid)
    if not typ:
        return {"aufwand": 0.0, "kosten": 0.0, "dauer": 0.0, "gewinn": 0.0,
                "kosten_vollstaendig": 1.0}
    gewinn = float(sum(f["gewinn"] for f in
                       handlungsfelder(state, uid, typ, types[typ], features, weights)))

    eigene = [m for m in massnahmen_liste(state, types, features, weights)
              if uid in m["einheiten"]]
    aufwand = n_gew = 0.0
    kosten_bez = k_gew = 0.0
    dauer_bez = d_gew = 0.0
    for m in eigene:
        anteil = 1.0 / len(m["einheiten"])
        n_gew += anteil
        aufwand += (m["aufwand"] or 0) * anteil
        if m["kosten"] is not None:
            kosten_bez += m["kosten"] * anteil
            k_gew += anteil
        if m["dauer"] is not None:
            dauer_bez += m["dauer"] * anteil
            d_gew += anteil
    kosten = kosten_bez * (n_gew / k_gew) if k_gew > 0 else 0.0
    dauer = dauer_bez * (n_gew / d_gew) if d_gew > 0 else 0.0
    return {
        "aufwand": aufwand,
        "kosten": kosten,
        "dauer": dauer,
        "gewinn": gewinn,
        "kosten_vollstaendig": (k_gew / n_gew) if n_gew > 0 else 1.0,
    }


# Standard-Kriterien der Nutzwertanalyse. 'richtung' min bedeutet: kleiner ist
# besser (frueher angehen); 'auto' verweist auf eine abgeleitete Kennzahl.
NWA_STANDARD = [
    {"name": "Aufwand", "gewicht": 25.0, "richtung": "min", "auto": "aufwand"},
    {"name": "Kosten", "gewicht": 25.0, "richtung": "min", "auto": "kosten"},
    {"name": "Risiko", "gewicht": 25.0, "richtung": "min", "auto": None},
    {"name": "Stückzahl", "gewicht": 25.0, "richtung": "max", "auto": None},
]


def nwa_kriterien(state):
    """Liste der Nutzwert-Kriterien, beim ersten Zugriff mit den Standard-Kriterien
    vorbelegt. Jedes Kriterium: name, gewicht, richtung ('min'/'max'), auto."""
    if "nwa_kriterien" not in state:
        state["nwa_kriterien"] = [dict(k) for k in NWA_STANDARD]
    return state["nwa_kriterien"]


def nwa_add_kriterium(state, name, richtung="min"):
    """Fuegt ein eigenes Kriterium hinzu (manuell zu bewerten, Startgewicht 0)."""
    name = (name or "").strip()
    if not name or any(k["name"] == name for k in nwa_kriterien(state)):
        return
    nwa_kriterien(state).append(
        {"name": name, "gewicht": 0.0, "richtung": richtung, "auto": None})


def nwa_remove_kriterium(state, name):
    """Entfernt ein Kriterium und die zugehoerigen manuellen Bewertungen."""
    state["nwa_kriterien"] = [k for k in nwa_kriterien(state) if k["name"] != name]
    w = state.get("nwa_werte", {})
    for key in [k for k in w if k[1] == name]:
        w.pop(key, None)


def nwa_set_gewicht(state, name, gewicht):
    for k in nwa_kriterien(state):
        if k["name"] == name:
            k["gewicht"] = float(gewicht)


def nwa_set_richtung(state, name, richtung):
    for k in nwa_kriterien(state):
        if k["name"] == name:
            k["richtung"] = richtung


def nwa_set_wert(state, uid, name, wert):
    """Setzt die manuelle (oder ueberschriebene) Bewertung einer Einheit fuer ein
    Kriterium. None loescht die manuelle Angabe (dann gilt wieder der auto-Wert)."""
    w = state.setdefault("nwa_werte", {})
    if wert is None:
        w.pop((uid, name), None)
    else:
        w[(uid, name)] = float(wert)


def nwa_get_wert(state, uid, kriterium, types, features, weights=None):
    """Rohwert einer Einheit fuer ein Kriterium: die manuelle Bewertung, sonst der
    abgeleitete auto-Wert (Aufwand/Kosten aus den Kennzahlen), sonst 0."""
    manuell = state.get("nwa_werte", {}).get((uid, kriterium["name"]))
    if manuell is not None:
        return manuell
    if kriterium["auto"]:
        return einheit_kennzahlen(state, uid, types, features, weights)[
            kriterium["auto"]]
    return 0.0


def nwa_nutzwerte(state, types, features, weights=None):
    """Nutzwert je Einheit nach Zangemeister: je Kriterium werden die Rohwerte ueber
    die Einheiten min-max-normiert und gepolt (bei richtung 'min' ist klein besser),
    mit dem normierten Gewicht multipliziert und summiert. Rueckgabe: Liste
    (uid, nutzwert), absteigend nach Nutzwert. Sind alle Werte eines Kriteriums
    gleich, traegt es fuer alle Einheiten gleich bei."""
    einheiten = [u for u in state["units"] if get_zieltyp(state, u)]
    if not einheiten:
        return []
    krit = nwa_kriterien(state)
    gew_summe = sum(k["gewicht"] for k in krit) or 1.0
    nutz = {u: 0.0 for u in einheiten}
    for k in krit:
        werte = {u: nwa_get_wert(state, u, k, types, features, weights)
                 for u in einheiten}
        lo, hi = min(werte.values()), max(werte.values())
        spanne = hi - lo
        for u in einheiten:
            if spanne == 0:
                norm = 1.0
            elif k["richtung"] == "min":
                norm = (hi - werte[u]) / spanne
            else:
                norm = (werte[u] - lo) / spanne
            nutz[u] += (k["gewicht"] / gew_summe) * norm
    return sorted(nutz.items(), key=lambda x: x[1], reverse=True)


def unternehmensweite_uebersicht(state, types, features, weights=None):
    """Zeilen der unternehmensweiten Handlungsfeld-Uebersicht, sortiert nach
    Merkmal. Eine zusammengelegte Kombination aus Merkmal und Ziel-Ausprägung
    erscheint als eine Zeile mit allen beteiligten Einheiten und dem summierten
    Aehnlichkeitsgewinn, alle uebrigen Handlungsfelder als eigene Zeile je Einheit.

    Jede Zeile ist ein dict mit merkmal, soll, einheiten (Liste von uids), ist
    (gemeinsamer Wert oder None bei Unterschied), aufwand (hoechster der Gruppe oder
    None), gewinn (Anteil, bei Buendelung summiert), gebuendelt (bool) und buendelbar
    (bool, also mehr als eine Einheit mit gleicher Kombination)."""
    reihenfolge = {m: i for i, m in enumerate(features)}
    from collections import OrderedDict
    gruppen = OrderedDict()
    for f in alle_handlungsfelder(state, types, features, weights):
        gruppen.setdefault((f["merkmal"], f["soll"]), []).append(f)

    zeilen = []
    for (merkmal, soll), gruppe in gruppen.items():
        buendelbar = len(gruppe) >= 2
        if buendelbar and ist_gebuendelt(state, merkmal, soll):
            zeilen.append({
                "merkmal": merkmal, "soll": soll,
                "einheiten": [f["uid"] for f in gruppe],
                "ist_werte": [f["ist"] for f in gruppe],   # parallel zu einheiten
                "aufwand": max((f["aufwand"] or 0) for f in gruppe) or None,
                "gewinn": sum(f["gewinn"] for f in gruppe),
                "gebuendelt": True, "buendelbar": True,
            })
        else:
            for f in gruppe:
                zeilen.append({
                    "merkmal": merkmal, "soll": soll,
                    "einheiten": [f["uid"]],
                    "ist_werte": [f["ist"]],
                    "aufwand": f["aufwand"],
                    "gewinn": f["gewinn"],
                    "gebuendelt": False, "buendelbar": buendelbar,
                })
    zeilen.sort(key=lambda z: (reihenfolge.get(z["merkmal"], 10 ** 9), str(z["soll"])))
    return zeilen


# Maßnahmen der unternehmensweiten Ebene. Ein Handlungsfeld wird hier eindeutig
# ueber einen Schluessel angesprochen, damit ein zusammengelegtes Vorhaben genau
# eine Massnahme traegt und nicht je Einheit eine eigene.
def hf_schluessel(zeile):
    """Eindeutiger, hashbarer Schluessel eines Handlungsfeldes der
    unternehmensweiten Ebene. Ein zusammengelegtes Handlungsfeld wird ueber Merkmal
    und Ziel bestimmt, ein einzelnes zusaetzlich ueber seine Einheit."""
    if zeile["gebuendelt"]:
        return ("b", zeile["merkmal"], zeile["soll"])
    return ("e", zeile["einheiten"][0], zeile["merkmal"])


def get_hf_massnahme(state, schluessel):
    """Massnahme eines unternehmensweiten Handlungsfeldes: dict mit text, phase, wer."""
    m = state.setdefault("hf_massnahmen", {})
    return m.get(schluessel, {"text": "", "phase": None, "wer": "", "voraussetzung": ""})


def set_hf_massnahme(state, schluessel, text=None, phase=None, wer=None,
                     voraussetzung=None):
    """Setzt Text, Phase, Verantwortlichkeit oder Voraussetzung eines
    unternehmensweiten Handlungsfeldes. Nur uebergebene Felder werden geaendert."""
    m = state.setdefault("hf_massnahmen", {})
    e = m.setdefault(schluessel, {"text": "", "phase": None, "wer": "", "voraussetzung": ""})
    if text is not None:
        e["text"] = text
    if phase is not None:
        e["phase"] = phase
    if wer is not None:
        e["wer"] = wer
    if voraussetzung is not None:
        e["voraussetzung"] = voraussetzung


def hf_phase_vorbelegen(state, zeilen):
    """Schlaegt fuer jedes Handlungsfeld ohne gesetzte Phase eine Phase anhand des
    Aufwands vor. Danach ist die Phase frei aenderbar. Idempotent."""
    for z in zeilen:
        k = hf_schluessel(z)
        if get_hf_massnahme(state, k)["phase"] is None:
            set_hf_massnahme(state, k, phase=(z["aufwand"] or PHASEN[-1]))


def hf_nach_phase(state, zeilen):
    """Gruppiert die unternehmensweiten Handlungsfelder nach ihrer Phase.
    Rueckgabe: dict {1: [zeile, ...], 2: [...], 3: [...]}."""
    gruppen = {p: [] for p in PHASEN}
    for z in zeilen:
        p = get_hf_massnahme(state, hf_schluessel(z))["phase"] or PHASEN[-1]
        gruppen.setdefault(p, []).append(z)
    return gruppen


def get_massnahme(state, uid, merkmal):
    """Massnahme zu einem Handlungsfeld: dict mit text, phase und wer.
    Fehlt ein Eintrag, werden leere Werte zurueckgegeben."""
    eintrag = state.get("massnahmen", {}).get(uid, {}).get(merkmal, {})
    return {"text": eintrag.get("text", ""),
            "phase": eintrag.get("phase"),
            "wer": eintrag.get("wer", "")}


def set_massnahme(state, uid, merkmal, text=None, phase=None, wer=None):
    """Setzt einzelne Felder einer Massnahme. Nicht uebergebene Felder bleiben
    unveraendert, sodass die drei Eingaben unabhaengig voneinander sind."""
    ebene = state.setdefault("massnahmen", {}).setdefault(uid, {})
    eintrag = ebene.setdefault(merkmal, {})
    if text is not None:
        eintrag["text"] = text
    if phase is not None:
        eintrag["phase"] = phase
    if wer is not None:
        eintrag["wer"] = wer


def phase_vorbelegen(state, uid, felder):
    """Schlaegt fuer jedes Handlungsfeld ohne gesetzte Phase eine Phase vor, und
    zwar zunaechst anhand seiner Aufwandsstufe. Der Anwender kann diese Zuordnung
    danach frei aendern, etwa wenn Abhaengigkeiten zwischen Massnahmen eine andere
    Reihenfolge erfordern. Ab diesem Zeitpunkt ist die Phase von der Aufwandsstufe
    unabhaengig. Idempotent."""
    for f in felder:
        if get_massnahme(state, uid, f["merkmal"])["phase"] is None:
            set_massnahme(state, uid, f["merkmal"],
                          phase=f["aufwand"] or PHASEN[-1])


def massnahmen_nach_phase(state, uid, felder):
    """Gruppiert die Handlungsfelder nach der gewaehlten Phase.
    Rueckgabe: dict {1: [feld, ...], 2: [...], 3: [...]}."""
    gruppen = {k: [] for k in PHASEN}
    for f in felder:
        phase = get_massnahme(state, uid, f["merkmal"])["phase"] or PHASEN[-1]
        gruppen.setdefault(phase, []).append(f)
    return gruppen
