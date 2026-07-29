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
        # WELT je Einheit und Merkmal (Tims Regel, explizit ueber Dropdown):
        # unit_id -> { merkmal -> "ist_bekannt" | "ist_unbekannt" }. Standard: ist_bekannt.
        "welt": {},
        # GEWICHTE je Merkmal (global, nicht je Einheit): merkmal -> float (>= 1).
        # Relative Wichtigkeit fuer die Aehnlichkeitsmasse; 1 = normal, 2 = doppelt,
        # usw. Wird beim Laden aus den Excel-Startwerten (falls vorhanden) oder mit
        # 1 initialisiert und in der Zustandserfassung editiert.
        "gewichte": {},
        # MASSNAHMEN je Einheit und Merkmal (nur fuer Handlungsfelder des Zieltyps):
        # unit_id -> { merkmal -> {"text": str, "etappe": 1|2|3, "wer": str} }.
        # Der Freitext und die Verantwortlichkeit stammen vollstaendig vom Anwender,
        # die Etappe ist aus dem Aufwand vorbelegt und aenderbar.
        "massnahmen": {},
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
    state["massnahmen"][uid] = {}   # Massnahmen je Merkmal (lazy angelegt)
    state["welt"][uid] = {m: WELT_IST_BEKANNT for m in features}
    return uid


def remove_unit(state, uid):
    """Entfernt eine Betrachtungseinheit vollstaendig (Ist, Potential,
    Ausschluss, engere Auswahl, Zieltyp, Soll, Aufwand, Welt). Der Zaehler
    'vergeben' bleibt unveraendert."""
    for ebene in ("matrix", "potential", "ausschluss", "engere_auswahl",
                  "zieltyp", "soll", "aufwand", "welt", "massnahmen"):
        if uid in state[ebene]:
            del state[ebene][uid]
    if uid in state["units"]:
        state["units"].remove(uid)


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
    """Setzt den Potential/Ausschluss-Status EINER Auspraegung.
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
    """Liefert den Status EINER Auspraegung: "", "potential" oder "ausschluss"."""
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
      - ZELLE_POTENTIAL / ZELLE_AUSSCHLUSS: setzt den PA-Status; falls die
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
    """Soll-Menge eines Merkmals: Ist-Auspraegung (falls echte Auspraegung)
    vereinigt mit den Potential-Auspraegungen.
    OFFEN traegt nichts bei (dann zaehlt nur das Potential)."""
    menge = set(potential_unit.get(merkmal, set()))
    ist = auswahl.get(merkmal, OFFEN)
    if ist is not OFFEN:
        menge.add(ist)
    return menge


def uebereinstimmung_soll(auswahl, potential_unit, typ_profil, features, weights=None, ausschluss_unit=None):
    """SOLL-KPI (SOLL-MINIMUM): gewichteter Soll-Uebereinstimmungsgrad mit EINEM Typ.

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
    """delta_i fuer den Soll-Zustand: True, wenn (Ist ∪ Potenzial) nicht leer ist."""
    return len(_soll_menge(auswahl, potential_unit, merkmal)) > 0


def _soll_treffer(auswahl, potential_unit, typ_profil, merkmal):
    """Lokale Soll-Aehnlichkeit: 1, wenn (Ist ∪ Potenzial) das Profil schneidet."""
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
    die Soll-Festlegung nimmt. Leer, falls keine gewaehlt."""
    return state["engere_auswahl"].get(uid, [])


def is_in_engere_auswahl(state, uid, typ_name):
    """True, wenn der Typ fuer die Einheit in der engeren Auswahl ist."""
    return typ_name in state["engere_auswahl"].get(uid, [])


def add_to_engere_auswahl(state, uid, typ_name):
    """Nimmt einen Typ in die engere Auswahl auf und legt seine (leere) eigene
    Soll-Konfiguration an. Gibt True bei Erfolg zurueck, False wenn die Auswahl
    den Typ bereits enthaelt."""
    liste = state["engere_auswahl"].setdefault(uid, [])
    if typ_name in liste:
        return False
    liste.append(typ_name)
    state["soll"].setdefault(uid, {}).setdefault(typ_name, {})
    return True


def remove_from_engere_auswahl(state, uid, typ_name):
    """Entfernt einen Typ aus der engeren Auswahl samt seiner Soll-Konfiguration
    und seiner Aufwaende. War er der finale Zieltyp, wird auch dieser zurueckgesetzt."""
    liste = state["engere_auswahl"].get(uid, [])
    if typ_name in liste:
        liste.remove(typ_name)
    state["soll"].get(uid, {}).pop(typ_name, None)
    state["aufwand"].get(uid, {}).pop(typ_name, None)
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
    """Setzt die angestrebte Soll-Auspraegung eines Merkmals fuer EINEN Typ
    (oder OFFEN). Jeder Typ der engeren Auswahl hat eine eigene Soll-Konfiguration.
    Aendert sich das Soll, wird der zugehoerige Aufwand verworfen, da er sich auf
    die zuvor gewaehlte Auspraegung bezog."""
    soll_typ = state["soll"].setdefault(uid, {}).setdefault(typ_name, {})
    if soll_typ.get(merkmal, OFFEN) != auspraegung:
        state["aufwand"].get(uid, {}).get(typ_name, {}).pop(merkmal, None)
    soll_typ[merkmal] = auspraegung


def get_soll(state, uid, typ_name, merkmal):
    """Liefert die angestrebte Soll-Auspraegung eines Merkmals fuer EINEN Typ,
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


def set_aufwand(state, uid, typ_name, merkmal, wert):
    """Setzt den geschaetzten Aufwand (1=gering, 2=mittel, 3=hoch), um das Merkmal
    vom heutigen Zustand in die gewaehlte Soll-Auspraegung zu bringen, fuer EINEN Typ."""
    state["aufwand"].setdefault(uid, {}).setdefault(typ_name, {})[merkmal] = wert


def get_aufwand(state, uid, typ_name, merkmal):
    """Liefert den geschaetzten Aufwand oder None (noch nicht geschaetzt)."""
    return state["aufwand"].get(uid, {}).get(typ_name, {}).get(merkmal)


def soll_kandidaten(state, uid, typ_profil, merkmal, status, optionen):
    """Auspraegungen mit Wahl-Knopf in der Soll-Festlegung, abhaengig vom Fall:
      - Fall 2 (Potenzial passt):          Ist + die Potenziale, die im Profil liegen.
      - Fall 3 (weder Ist noch Potenzial): Ist + alle Profilauspraegungen.
      - Fall 4 & 5 (kein Ist):             alle Auspraegungen (freie Wahl).
      - Fall 1 (Ist passt):                keine (Soll ist automatisch das Ist).
    In Fall 2 werden bewusst nur die als Potenzial erfassten, im Profil liegenden
    Auspraegungen angeboten (nicht jede Profilauspraegung), da hier bereits ein
    erreichbarer, passender Weg erfasst ist. In Fall 3 gibt es keinen solchen Weg,
    daher stehen alle Profilauspraegungen zur Wahl. Rueckgabe als Menge."""
    profil = typ_profil.get(merkmal, set())
    ist = get_choice(state, uid, merkmal)
    if status == STATUS_POTENTIAL:
        kand = set(get_potential(state, uid, merkmal) & profil)
        if ist is not OFFEN:
            kand.add(ist)
        return kand
    if status == STATUS_IST_UNPASSEND:
        kand = set(profil)
        if ist is not OFFEN:
            kand.add(ist)
        return kand
    if status in (STATUS_OFFEN, STATUS_BLOCKIERT):
        return set(optionen)
    return set()


def soll_vorbelegen(state, uid, typ_name, features):
    """Belegt das Soll mit dem Ist vor, wo es ein Ist gibt (Faelle 1, 2, 3) und
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
    """True, wenn fuer JEDES Merkmal eine Entscheidung vorliegt: ein Soll (eine
    Auspraegung oder bewusst NICHTS_ANSTREBEN) und, falls das Soll vom Ist abweicht,
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
    das Soll ueberall vorbelegt, damit nicht geoeffnete Tabs nicht blockieren."""
    hat_auswahl = False
    for uid in state["units"]:
        for typ in get_engere_auswahl(state, uid):
            hat_auswahl = True
            soll_vorbelegen(state, uid, typ, features)
            if not detaillierung_vollstaendig(state, uid, typ, features):
                return False
    return hat_auswahl


def soll_score_gestaffelt(state, uid, typ_name, typ_profil, features, weights=None):
    """Soll-Score je Aufwandsstufe fuer EINEN Typ, als dict {0: .., 1: .., 2: .., 3: ..}.

    Stufe 0 ist die bereits ohne jede Aenderung erreichte Uebereinstimmung.

    Ein Merkmal ist ein Treffer, wenn seine Soll-Auspraegung im Profil liegt. Der
    Erreichungsaufwand ist 0, wenn das Soll bereits dem Ist entspricht (schon
    erreicht, kein Aufwand), sonst der erfasste Aufwand (1..3). Ein Treffer zaehlt
    in Stufe k, wenn sein Erreichungsaufwand <= k. Der Score bei Stufe k ist damit
    die erreichte gewichtete Profilkonformitaet, wenn alle Aenderungen bis
    einschliesslich Aufwand k umgesetzt werden. Stufe 3 ist der volle Soll-Score.

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
    profilkonformen Annaeherung fuehren (Treffer mit Soll != Ist). Rueckgabe:
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
    """Bestimmt den Status EINES Merkmals der Einheit gegenueber EINEM Typ.
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
# konkrete Massnahme, die Etappe und die Verantwortlichkeit.

def handlungsfelder(state, uid, typ_name, typ_profil, features, weights=None):
    """Liste der Handlungsfelder eines Zieltyps. Ein Handlungsfeld erfuellt ZWEI
    Bedingungen: die Soll-Auspraegung weicht vom Ist ab (es gibt eine Veraenderung)
    UND sie liegt im Profil des Zieltyps (die Veraenderung fuehrt naeher an den Typ
    heran). Nur dann dient die Veraenderung der Annaeherung an den Zieltyp.

    Kein Handlungsfeld sind daher: Merkmale mit Soll = Ist (keine Veraenderung),
    Merkmale ohne angestrebte Ausprägung (bewusst nichts) und Merkmale mit
    profilfremder Soll-Ausprägung. Die beiden letzten sind bewusste Abweichungen
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


def nicht_erreichte_merkmale(state, uid, typ_name, typ_profil, features):
    """Merkmale, in denen der Zieltyp dauerhaft nicht erreicht wird, weil das Soll
    nicht im Profil liegt oder bewusst nichts angestrebt wurde. Sie begruenden,
    warum die erreichbare Ähnlichkeit unter eins bleibt."""
    offen = []
    for m in features:
        soll = get_soll(state, uid, typ_name, m)
        if soll == NICHTS_ANSTREBEN:
            offen.append((m, "bewusst nichts angestrebt"))
        elif soll is OFFEN:
            offen.append((m, "keine Soll-Angabe"))
        elif soll not in typ_profil.get(m, set()):
            offen.append((m, "Soll liegt ausserhalb des Profils"))
    return offen


def get_massnahme(state, uid, merkmal):
    """Massnahme zu einem Handlungsfeld: dict mit text, etappe und wer.
    Fehlt ein Eintrag, werden leere Werte zurueckgegeben."""
    eintrag = state.get("massnahmen", {}).get(uid, {}).get(merkmal, {})
    return {"text": eintrag.get("text", ""),
            "etappe": eintrag.get("etappe"),
            "wer": eintrag.get("wer", "")}


def set_massnahme(state, uid, merkmal, text=None, etappe=None, wer=None):
    """Setzt einzelne Felder einer Massnahme. Nicht uebergebene Felder bleiben
    unveraendert, sodass die drei Eingaben unabhaengig voneinander sind."""
    ebene = state.setdefault("massnahmen", {}).setdefault(uid, {})
    eintrag = ebene.setdefault(merkmal, {})
    if text is not None:
        eintrag["text"] = text
    if etappe is not None:
        eintrag["etappe"] = etappe
    if wer is not None:
        eintrag["wer"] = wer


def etappe_vorbelegen(state, uid, felder):
    """Belegt die Etappe jedes Handlungsfeldes mit seiner Aufwandsstufe vor, sofern
    noch keine gesetzt ist. Die Reihenfolge der Umsetzung folgt damit zunaechst dem
    Aufwand, kann vom Anwender aber geaendert werden, etwa wenn Abhaengigkeiten
    zwischen Massnahmen bestehen. Idempotent."""
    for f in felder:
        if get_massnahme(state, uid, f["merkmal"])["etappe"] is None:
            set_massnahme(state, uid, f["merkmal"],
                          etappe=f["aufwand"] or AUFWAND_HOCH)


def massnahmen_nach_etappe(state, uid, felder):
    """Gruppiert die Handlungsfelder nach der gewaehlten Etappe.
    Rueckgabe: dict {1: [feld, ...], 2: [...], 3: [...]}."""
    gruppen = {k: [] for k in AUFWAND_STUFEN}
    for f in felder:
        etappe = get_massnahme(state, uid, f["merkmal"])["etappe"] or AUFWAND_HOCH
        gruppen.setdefault(etappe, []).append(f)
    return gruppen