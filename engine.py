"""Logique métier : correspondance des colonnes Excel par nom d'en-tête,
détection collectif/individuel, regroupement en unités, et calcul des
valeurs (formatage identique à modGenererRapports.bas : FormatVal,
FormatValM2, BuildLocalisation, BuildRemarques).

Les colonnes sont retrouvées par NOM D'EN-TÊTE plutôt que par position fixe :
plus robuste que la macro VBA d'origine, dont les constantes COL_* supposent
que toutes les missions utilisent exactement la même disposition de colonnes
(ce qui n'est pas le cas : les fichiers "Commu" et "Studio" réels fournis par
Georges n'ont pas les mêmes colonnes aux mêmes positions).
"""
import re
import unicodedata
from datetime import datetime, date


def normalize_header(s) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.lower()
    s = re.sub(r"[°.,]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


FIELD_ALIASES = {
    "typeLogement": ["type logement"],
    "codePostal": ["code postal"],
    "localite": ["localite"],
    "adresse": ["adresse"],
    "numero": ["n"],
    "bte": ["bte"],
    "datePassage": ["date passage"],
    "dateCloture": ["date de cloture"],
    "etage": ["etage"],
    "avAr": ["av/ ar", "av/ar"],
    "gcd": ["g/c/d"],
    "nbOccupants": ["nbre occupants", "nbre occup"],
    "nbPiecesInd": ["nbre pieces individuels"],
    "nbPiecesColl": ["nbre pieces collectives"],
    "nbSdb": ["nbre sdb/sdd"],
    "nbWc": ["nbre wc"],
    "surfStudio": ["surface studios"],
    "surfPiecesInd": ["surface pieces individuelles (chambres)", "surface pieces individuelles"],
    "supCollective": ["superficie pieces habitables collectives", "superficie pieces collectives"],
    "supTotale": ["superficie totale commu et chambres", "superficie totale"],
    "remVent": ["rem ventilation"],
    "remHum": ["rem humidite"],
    "remElec": ["rem electricite"],
    "remDiv": ["rem divers"],
    "remGc": ["rem garde-corps", "rem garde corps"],
    "hsp": ["hsp"],
    "chassisM2": ["chassis m²", "chassis m2"],
    "hauteurAllege": ["hauteur d'allege"],
    "detect": ["detect"],
    # Identification du bailleur (colonnes pas encore présentes dans les
    # fichiers Excel de Georges au 22/09/2026 : ces alias sont prêts à
    # reconnaître ces colonnes dès qu'il les ajoutera, sans nouvelle
    # intervention de ma part).
    "nomBailleur": ["nom bailleur"],
    "prenomBailleur": ["prenom bailleur"],
    "adresseBailleur": ["adresse bailleur"],
    "numeroBailleur": ["n bailleur", "numero bailleur"],
    "cpBailleur": ["cp bailleur", "code postal bailleur"],
    "localiteBailleur": ["localite bailleur"],
    "telBailleur": ["telephone bailleur", "tel bailleur"],
    "mailBailleur": ["mail bailleur", "email bailleur"],
}

# Coordonnées de l'enquêteur (Georges) : fixes, toujours identiques d'une
# mission à l'autre (confirmé par Georges le 22/09/2026). Modifiables ici
# directement si elles changent un jour ; peut évoluer vers un petit écran
# de réglages dans l'application si Georges préfère les modifier lui-même
# sans dépendre d'une nouvelle version publiée.
ENQUETEUR = {
    "Nom expert": "BASTIAN",
    "Prénom expert": "Georges",
    "Numéro d'agrément ": "23/1273P",
    "Adresse expert ": "Avenue Reine Astrid",
    "N° expert": "16",
    "CP expert": "1300",
    "Localité expert": "Wavre",
    "Mail expert": "info@expertsnicolai.be",
    "Téléphone expert": "010/45.10.05",
}

# Petits tableaux verticaux "étiquette / valeur" (comme ceux que Georges a
# construits : "Identification du bailleur" et "Identification expert
# agréé", chacun avec une ligne par champ) — une seule fois par fichier
# Excel, où qu'ils se trouvent dedans (n'importe quelle feuille, n'importe
# quelle position), plutôt que répétés en colonne sur chaque ligne du
# tableau principal.
BAILLEUR_LABELS = {
    "nom bailleur": "Nom bailleur",
    "prenom bailleur": "Prénom bailleur",
    "adresse bailleur": "Adresse bailleur",
    "numero bailleur": "Numéro bailleur",
    "n bailleur": "Numéro bailleur",
    "cp bailleur": "CP bailleur",
    "localite bailleur": "Localité bailleur",
    "telephone bailleur": "Téléphone bailleur",
    "tel bailleur": "Téléphone bailleur",
    "mail bailleur": "Mail bailleur",
    "email bailleur": "Mail bailleur",
}

EXPERT_LABELS = {
    "nom expert": "Nom expert",
    "prenom expert": "Prénom expert",
    "adresse expert": "Adresse expert ",
    "n expert": "N° expert",
    "numero expert": "N° expert",
    "cp expert": "CP expert",
    "localite expert": "Localité expert",
    "telephone expert": "Téléphone expert",
    "tel expert": "Téléphone expert",
    "mail expert": "Mail expert",
    "email expert": "Mail expert",
    "numero d'agrement": "Numéro d'agrément ",
    "numero agrement": "Numéro d'agrément ",
}


# Certains classeurs (ex. "Ladeuze 20 - Studio Officiel.xlsx", 23/09/2026)
# n'utilisent pas d'étiquette qualifiée ("Nom bailleur") pour le tableau du
# bailleur, mais une étiquette "nue" ("Nom", "Prénom", "Adresse", "N°", "CP",
# "Téléphone", "Mail") placée sous un intitulé de section ("Identification du
# bailleur"). Le tableau de l'expert, dans le même classeur, utilise lui les
# étiquettes qualifiées habituelles ("Nom expert", etc.) — les deux formes
# doivent donc être reconnues, potentiellement dans le même fichier.
BASE_LABELS_BAILLEUR = {
    "nom": "Nom bailleur",
    "prenom": "Prénom bailleur",
    "adresse": "Adresse bailleur",
    "n": "Numéro bailleur",
    "numero": "Numéro bailleur",
    "cp": "CP bailleur",
    "code postal": "CP bailleur",
    "localite": "Localité bailleur",
    "telephone": "Téléphone bailleur",
    "tel": "Téléphone bailleur",
    "mail": "Mail bailleur",
    "email": "Mail bailleur",
}

BASE_LABELS_EXPERT = {
    "nom": "Nom expert",
    "prenom": "Prénom expert",
    "adresse": "Adresse expert ",
    "n": "N° expert",
    "numero": "N° expert",
    "cp": "CP expert",
    "code postal": "CP expert",
    "localite": "Localité expert",
    "telephone": "Téléphone expert",
    "tel": "Téléphone expert",
    "mail": "Mail expert",
    "email": "Mail expert",
}

# Ensemble des étiquettes "nues" reconnues (utilisé pour repérer les lignes
# d'intitulé de section : une ligne comme "Identification du bailleur" n'est
# PAS elle-même une de ces étiquettes nues, contrairement à "Nom" ou "CP").
_BASE_LABEL_WORDS = set(BASE_LABELS_BAILLEUR) | set(BASE_LABELS_EXPERT)


def _first_value_right_of(row, from_index):
    for c2 in range(from_index + 1, len(row)):
        v = row[c2]
        if v is not None and str(v).strip() != "":
            return v
    return None


def scan_bailleur_expert(all_sheets_rows):
    """Comme scan_labeled_values, mais reconnaît EN PLUS les tableaux
    "étiquette / valeur" du bailleur et/ou de l'expert qui utilisent des
    étiquettes "nues" ("Nom", "Prénom", "Adresse", "N°", "CP", "Téléphone",
    "Mail") placées sous un intitulé de section ("Identification du
    bailleur", "Identification expert agréé" — toléré même avec une coquille
    du type "Idenfication"), en plus des étiquettes déjà qualifiées ("Nom
    bailleur", "Nom expert", etc., reconnues comme avant, indépendamment de
    toute section).

    Le repérage d'un intitulé de section ne dépend d'aucune orthographe
    exacte : toute cellule dont le texte normalisé contient "bailleur" (ou
    "expert") sans être elle-même une étiquette nue ou qualifiée connue
    ("nom bailleur", "cp", etc.) est traitée comme un début de section. La
    section reste active jusqu'à la suivante, feuille par feuille. Ne lève
    jamais d'exception : un tableau absent ou mal placé donne simplement un
    dict vide pour les champs correspondants (comportement inchangé pour les
    fichiers qui n'ont pas encore ce tableau)."""
    combined_qualified = {**BAILLEUR_LABELS, **EXPERT_LABELS}
    found = {}
    try:
        for rows in all_sheets_rows:
            section = None
            for row in rows:
                first_idx = None
                first_norm = ""
                for c, v in enumerate(row):
                    norm = normalize_header(v)
                    if norm:
                        first_idx = c
                        first_norm = norm
                        break
                if first_idx is None:
                    continue

                if first_norm in combined_qualified:
                    tag = combined_qualified[first_norm]
                    if tag not in found:
                        value = _first_value_right_of(row, first_idx)
                        if value is not None:
                            found[tag] = format_val(value)
                    section = "bailleur" if first_norm in BAILLEUR_LABELS else "expert"
                    continue

                if "bailleur" in first_norm:
                    remainder = first_norm.replace("bailleur", "", 1).strip()
                    if remainder not in _BASE_LABEL_WORDS:
                        section = "bailleur"
                        continue
                if "expert" in first_norm:
                    remainder = first_norm.replace("expert", "", 1).strip()
                    if remainder not in _BASE_LABEL_WORDS:
                        section = "expert"
                        continue

                if section and first_norm in _BASE_LABEL_WORDS:
                    label_dict = BASE_LABELS_BAILLEUR if section == "bailleur" else BASE_LABELS_EXPERT
                    tag = label_dict.get(first_norm)
                    if tag and tag not in found:
                        value = _first_value_right_of(row, first_idx)
                        if value is not None:
                            found[tag] = format_val(value)
    except Exception:
        pass
    return found


def scan_labeled_values(all_sheets_rows, label_map):
    """Cherche, dans toutes les lignes de toutes les feuilles fournies, une
    cellule dont le texte correspond (après normalisation) à une étiquette
    connue de label_map, et prend comme valeur la première cellule non vide
    à sa droite sur la même ligne. Retourne {balise: valeur_formatee}. Ne
    lève jamais d'exception : un tableau absent ou mal placé donne
    simplement un dict vide (comportement inchangé pour les fichiers qui
    n'ont pas encore ce tableau)."""
    found = {}
    try:
        for rows in all_sheets_rows:
            for row in rows:
                n = len(row)
                for c in range(n):
                    label_norm = normalize_header(row[c])
                    if not label_norm or label_norm not in label_map:
                        continue
                    tag = label_map[label_norm]
                    if tag in found:
                        continue  # on garde la première occurrence trouvée
                    value = None
                    for c2 in range(c + 1, n):
                        v = row[c2]
                        if v is not None and str(v).strip() != "":
                            value = v
                            break
                    if value is not None:
                        found[tag] = format_val(value)
    except Exception:
        pass
    return found


REQUIRED_FIELDS = ["typeLogement", "codePostal", "localite", "adresse", "numero", "bte", "datePassage"]


def build_header_map(header_row):
    normalized = [normalize_header(h) for h in header_row]
    m = {}
    used = set()
    for field, aliases in FIELD_ALIASES.items():
        found = None
        for alias in aliases:
            if alias in normalized:
                found = normalized.index(alias)
                break
        if found is not None:
            m[field] = found
            used.add(found)
    unmatched = [header_row[i] for i, h in enumerate(normalized) if h and i not in used]
    missing = [f for f in REQUIRED_FIELDS if f not in m]
    return m, unmatched, missing


def cell(row, m, field):
    idx = m.get(field)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def is_empty(v) -> bool:
    return v is None or v == ""


def format_val(v) -> str:
    if is_empty(v):
        return ""
    if isinstance(v, (datetime, date)):
        return f"{v.day:02d}/{v.month:02d}/{v.year:04d}"
    if isinstance(v, bool):
        s = str(v)
    elif isinstance(v, (int, float)):
        rounded = round(float(v), 2)
        if rounded == int(rounded):
            s = str(int(rounded))
        else:
            s = f"{rounded:.2f}".rstrip("0").rstrip(".")
            s = s.replace(".", ",")
    else:
        s = str(v)
    s = re.sub(r"[,.]+$", "", s)
    return s.strip()


def format_val_m2(v) -> str:
    s = format_val(v)
    return f"{s} m²" if s else ""


def to_num(v) -> float:
    if is_empty(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return 0.0


def build_localisation(row, m) -> str:
    parts = []
    for f in ("etage", "avAr", "gcd"):
        v = cell(row, m, f)
        if not is_empty(v) and str(v).strip():
            parts.append(str(v).strip())
    return ", ".join(parts)


def build_remarques(row, m) -> str:
    labels = [
        ("remVent", "Ventilation"),
        ("remHum", "Humidité"),
        ("remElec", "Electricité"),
        ("remDiv", "Divers"),
        ("remGc", "Garde-corps"),
    ]
    parts = []
    for f, label in labels:
        v = cell(row, m, f)
        if not is_empty(v) and str(v).strip():
            parts.append(f"{label} : {str(v).strip()}")
    return " ; ".join(parts)


def detect_format(rows, m, header_row_index=0) -> str:
    for row in rows[header_row_index + 1:]:
        t = normalize_header(cell(row, m, "typeLogement"))
        if t[:7] == "chambre" or t[:5] == "commu":
            return "collectif"
    for row in rows[header_row_index + 1:]:
        v = cell(row, m, "typeLogement")
        if not is_empty(v) and str(v).strip():
            return "individuel"
    return ""


def build_units_collectif(rows, m, header_row_index=0):
    units = []
    data_rows = rows[header_row_index + 1:]
    i = 0
    n = len(data_rows)
    while i < n:
        row = data_rows[i]
        t = normalize_header(cell(row, m, "typeLogement"))
        if t[:5] == "commu":
            chambres = []
            j = i + 1
            while j < n:
                t2 = normalize_header(cell(data_rows[j], m, "typeLogement"))
                if t2[:7] != "chambre":
                    break
                chambres.append(data_rows[j])
                j += 1
            units.append({"format": "collectif", "commun_row": row, "chambres": chambres})
            i = j
        else:
            i += 1
    return units


def build_units_individuel(rows, m, header_row_index=0):
    units = []
    for row in rows[header_row_index + 1:]:
        v = cell(row, m, "typeLogement")
        if not is_empty(v) and str(v).strip():
            units.append({"format": "individuel", "row": row})
    return units


def clean_file_name(s: str) -> str:
    s = str(s)
    for ch in ("/", "\\", ":"):
        s = s.replace(ch, "-")
    s = re.sub(r"\s+", "_", s)
    return s


def unit_label(unit, m):
    row = unit["commun_row"] if unit["format"] == "collectif" else unit["row"]
    adresse = str(cell(row, m, "adresse") or "").strip()
    n = format_val(cell(row, m, "numero"))
    bte_val = cell(row, m, "bte")
    bte = str(bte_val).strip() if bte_val is not None else ""
    file_base = clean_file_name(f"Rapport_{adresse}_{n}_{bte}")
    return {"adresse": adresse, "n": n, "bte": bte, "file_base": file_base}


def detect_collectif(unit, m) -> str:
    """Colonne « Détect » d'un logement collectif : valeur de la ligne
    « Communautaire » ; si elle est vide, valeur(s) des chambres (une seule
    si elles sont toutes identiques, sinon les différentes valeurs séparées
    par « / »)."""
    v = str(cell(unit["commun_row"], m, "detect") or "").strip()
    if v:
        return v
    vals = []
    for c in unit["chambres"]:
        x = str(cell(c, m, "detect") or "").strip()
        if x and x not in vals:
            vals.append(x)
    return " / ".join(vals)


def build_data_spec_collectif(unit, m, extra_scalars=None):
    row = unit["commun_row"]
    chambres = unit["chambres"]
    nb_chambres = len(chambres)

    nb_sanit = to_num(cell(row, m, "nbSdb")) + to_num(cell(row, m, "nbWc"))
    localite = str(cell(row, m, "localite") or "").strip()
    date_passage_val = format_val(cell(row, m, "datePassage"))
    date_cloture_val = format_val(cell(row, m, "dateCloture"))
    unit_bte_val = str(cell(row, m, "bte")).strip() if cell(row, m, "bte") is not None else ""

    scalars = {
        "Adresse": str(cell(row, m, "adresse") or "").strip(),
        "N": format_val(cell(row, m, "numero")),
        "N°": format_val(cell(row, m, "numero")),
        # NB : "Bte" n'est PAS un scalaire ici (voir positional ci-dessous) :
        # ce même nom de balise est aussi utilisé, dans les gabarits collectifs
        # mis à jour par Georges le 23/09/2026, à l'intérieur du tableau
        # répété par chambre — il faut donc une valeur différente à chaque
        # occurrence, ce qu'un scalaire ne permet pas (un scalaire s'applique
        # identiquement à toutes les occurrences de la balise).
        "Code_postal": format_val(cell(row, m, "codePostal")),
        "Localité": localite,
        "Localite": localite,  # alias, au cas où un gabarit utiliserait l'orthographe sans accent
        "Date_de_passage": date_passage_val,
        "Date passage": date_passage_val,  # alias, tag utilisé ailleurs dans certains gabarits
        "Date de clôture": date_cloture_val,
        "Dâte de clôture": date_cloture_val,  # alias, coquille présente dans un gabarit
        "Etage": str(cell(row, m, "etage") or "").strip(),
        "Détect": detect_collectif(unit, m),
        "Rem ventilation": str(cell(row, m, "remVent") or "").strip(),
        "Rem Humidité": str(cell(row, m, "remHum") or "").strip(),
        "Rem Electricité": str(cell(row, m, "remElec") or "").strip(),
        "Rem divers": str(cell(row, m, "remDiv") or "").strip(),
        "Rem garde-corps": str(cell(row, m, "remGc") or "").strip(),
        "Nbre_sanitaires": format_val(nb_sanit),
        "Nbre_toilettes": format_val(cell(row, m, "nbWc")),
        "Nbre SDB/SDD": format_val(cell(row, m, "nbSdb")),
        "Nbre WC": format_val(cell(row, m, "nbWc")),
        "Remarques": build_remarques(row, m),
        "Localisation_precise": build_localisation(row, m),
        "Nbre_pieces_individuelles": str(nb_chambres),
        "Nbre_pieces_collectives": format_val(cell(row, m, "nbPiecesColl")),
        "Nbre pièces collectives": format_val(cell(row, m, "nbPiecesColl")),
        "Capacite_accueil": "",
        # Identification du bailleur (vide tant que Georges n'a pas ajouté
        # les colonnes correspondantes dans son Excel).
        "Nom bailleur": str(cell(row, m, "nomBailleur") or "").strip(),
        "Prénom bailleur": str(cell(row, m, "prenomBailleur") or "").strip(),
        "Adresse bailleur": str(cell(row, m, "adresseBailleur") or "").strip(),
        "Numéro bailleur": format_val(cell(row, m, "numeroBailleur")),
        "CP bailleur": format_val(cell(row, m, "cpBailleur")),
        "Localité bailleur": str(cell(row, m, "localiteBailleur") or "").strip(),
        "Téléphone bailleur": str(cell(row, m, "telBailleur") or "").strip(),
        "Mail bailleur": str(cell(row, m, "mailBailleur") or "").strip(),
    }
    scalars.update(ENQUETEUR)
    # Tableau "étiquette / valeur" trouvé dans le classeur (bailleur/expert) :
    # prioritaire sur les valeurs par défaut ci-dessus, puisque c'est la
    # source d'information la plus à jour choisie par Georges.
    if extra_scalars:
        scalars.update(extra_scalars)

    # Le tableau "par chambre" des gabarits réserve toujours 19 lignes
    # physiques (MAX_CHAMBRES_TEMPLATE), quel que soit le nombre réel de
    # chambres de ce logement : on complète donc à 19 avec des valeurs
    # vides avant d'ajouter d'éventuelles valeurs supplémentaires utilisées
    # ailleurs dans le gabarit (voir Bte et Nbre occupants ci-dessous), pour
    # que ces valeurs supplémentaires tombent bien sur la bonne occurrence
    # de la balise et non sur une ligne de chambre inutilisée.
    pad = [""] * max(0, MAX_CHAMBRES_TEMPLATE - nb_chambres)

    sup_coll = format_val_m2(cell(row, m, "supCollective"))
    nb_pieces_ind = [format_val(cell(c, m, "nbPiecesInd")) for c in chambres]
    surf_ind = [format_val_m2(cell(c, m, "surfPiecesInd")) for c in chambres]
    sup_totale = [format_val_m2(cell(c, m, "supTotale")) for c in chambres]
    av_ar = [str(cell(c, m, "avAr") or "").strip() for c in chambres]
    gcd = [str(cell(c, m, "gcd") or "").strip() for c in chambres]
    total_occupants = format_val(sum(to_num(cell(c, m, "nbOccupants")) for c in chambres))

    positional = {
        "N_chambre": [format_val(cell(c, m, "bte")) for c in chambres],
        "Nbre_pieces_individuels": nb_pieces_ind,
        "Nbre pièces individuels": nb_pieces_ind,
        "Surface_chambre": surf_ind,
        "Surface pièces individuelles": surf_ind,
        "Superficie_totale": sup_totale,
        "Superficie totale": sup_totale,
        "Superficie_collective": [sup_coll] * (nb_chambres + 1),
        "Superficie pièces  collectives": [sup_coll] * (nb_chambres + 1),
        "Av/ AR": av_ar,
        "G/C/D": gcd,
        # "Bte" par chambre (table), encadré par la ou les deux occurrences
        # "scalaires" historiques (identification du logement, avant le
        # tableau ; et, pour les gabarits 1B/1D uniquement, une seconde fois
        # dans la section volet 2 après le tableau). Les gabarits qui n'ont
        # pas cette seconde occurrence (1A/1C) ignorent simplement la
        # dernière valeur du tableau ci-dessous.
        "Bte": [unit_bte_val] + [format_val(cell(c, m, "bte")) for c in chambres] + pad + [unit_bte_val],
        # "Nbre occupants" par chambre (table), puis la même valeur totale
        # répétée pour les quelques mentions "… ménage(s) / … personne(s)"
        # qui suivent le tableau (jusqu'à 8 occurrences couvertes, largement
        # suffisant ; les occurrences en trop, s'il y en a moins, sont
        # simplement ignorées).
        "Nbre occupants": [format_val(cell(c, m, "nbOccupants")) for c in chambres] + pad + [total_occupants] * 8,
    }
    return scalars, positional


def build_data_spec_individuel(unit, m, extra_scalars=None):
    row = unit["row"]
    localite = str(cell(row, m, "localite") or "").strip()
    date_passage_val = format_val(cell(row, m, "datePassage"))
    date_cloture_val = format_val(cell(row, m, "dateCloture"))
    surf_studio_val = format_val(cell(row, m, "surfStudio"))
    scalars = {
        "Adresse": str(cell(row, m, "adresse") or "").strip(),
        "N": format_val(cell(row, m, "numero")),
        "N°": format_val(cell(row, m, "numero")),
        "Bte": str(cell(row, m, "bte")).strip() if cell(row, m, "bte") is not None else "",
        "Code_postal": format_val(cell(row, m, "codePostal")),
        "Localité": localite,
        "Localite": localite,
        "Date_de_passage": date_passage_val,
        "Date passage": date_passage_val,
        "Date de clôture": date_cloture_val,
        "Dâte de clôture": date_cloture_val,
        "Etage": str(cell(row, m, "etage") or "").strip(),
        "Av/ AR": str(cell(row, m, "avAr") or "").strip(),
        "G/C/D": str(cell(row, m, "gcd") or "").strip(),
        "HSP": format_val(cell(row, m, "hsp")),
        "Châssis m²": format_val(cell(row, m, "chassisM2")),
        "Hauteur d'allège": format_val(cell(row, m, "hauteurAllege")),
        "Détect": str(cell(row, m, "detect") or "").strip(),
        "Rem ventilation": str(cell(row, m, "remVent") or "").strip(),
        "Rem Humidité": str(cell(row, m, "remHum") or "").strip(),
        "Rem Electricité": str(cell(row, m, "remElec") or "").strip(),
        "Rem divers": str(cell(row, m, "remDiv") or "").strip(),
        "Rem garde-corps": str(cell(row, m, "remGc") or "").strip(),
        "Remarques": build_remarques(row, m),
        "Localisation_precise": build_localisation(row, m),
        "Surface_studio": surf_studio_val,
        "Surface studios": surf_studio_val,
        "Nbre_occupants": format_val(cell(row, m, "nbOccupants")),
        "Nbre_pieces_individuelles": "",
        "Nbre_pieces_collectives": "",
        "Nbre_sanitaires": "",
        # Identification du bailleur (vide tant que Georges n'a pas ajouté
        # les colonnes correspondantes dans son Excel).
        "Nom bailleur": str(cell(row, m, "nomBailleur") or "").strip(),
        "Prénom bailleur": str(cell(row, m, "prenomBailleur") or "").strip(),
        "Adresse bailleur": str(cell(row, m, "adresseBailleur") or "").strip(),
        "Numéro bailleur": format_val(cell(row, m, "numeroBailleur")),
        "CP bailleur": format_val(cell(row, m, "cpBailleur")),
        "Localité bailleur": str(cell(row, m, "localiteBailleur") or "").strip(),
        "Téléphone bailleur": str(cell(row, m, "telBailleur") or "").strip(),
        "Mail bailleur": str(cell(row, m, "mailBailleur") or "").strip(),
    }
    scalars.update(ENQUETEUR)
    if extra_scalars:
        scalars.update(extra_scalars)
    return scalars, {}


MAX_CHAMBRES_TEMPLATE = 19
