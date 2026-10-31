"""État des lieux de SORTIE — saisie assistée, 100 % locale, sans IA.

Georges lit lui-même le document « valorisations manuscrites » et saisit les
valeurs dans l'onglet. Ce module :
  1. remplit une COPIE du tableau Excel type (« ELLS Détails Valorisation
     dégâts ») sans jamais supprimer de ligne ni toucher à la mise en page ;
  2. calcule le total exactement comme les formules du tableau ;
  3. remplit le PV Word (modèle du dossier 2026.2246) : parties, adresse,
     date de sortie, montant en chiffres et en toutes lettres majuscules,
     date de clôture et n° de dossier ;
  4. enregistre la saisie (fichier .json dans le dossier) pour pouvoir la
     rouvrir et la compléter plus tard.

Aucune connexion internet, aucune clé API, aucun coût.
"""
import ast
import json
import operator
import os
import re
import shutil
from datetime import date, datetime

import openpyxl

import docx_fill
import montant_lettres
import securite
import signature

APP_DIR = os.path.dirname(os.path.abspath(__file__))
GABARIT_DIR = os.path.join(APP_DIR, "gabarit_sortie")

SHEET_NAME = "Feuil1"

# --- Plan du tableau Excel type (gabarit vierge, 139 lignes) ---------------
CLES_ROWS = {"copies": 42, "cyl_std": 43, "cyl_secu": 44, "detecteurs": 45}
ENTRETIEN_ROWS = {
    "boiler": 52, "ramonage": 53, "chauffage": 54, "chauffage_gaz": 54, "adoucisseur": 55,
    "alarme": 56, "ventilation": 57, "fosse": 58,
}
ENTRETIEN_LABELS = {
    "boiler": "Entretien boiler",
    "ramonage": "Ramonage cheminée",
    "chauffage": "Entretien chaudière mazout",
    "chauffage_gaz": "Entretien chaudière gaz",
    "adoucisseur": "Entretien adoucisseur",
    "alarme": "Entretien alarme",
    "ventilation": "Entretien ventilation / filtre / double flux",
    "fosse": "Vidange fosse septique (2 m³)",
}
CLES_LABELS = {
    "copies": "Copies clés à prévoir",
    "cyl_std": "Remplacement cylindre standard",
    "cyl_secu": "Remplacement cylindre sécurité",
    "detecteurs": "Détecteur manquant / pile HS",
}
# Tarifs HTVA (décisions GB du 31/10/2026 : mazout 250 €, gaz 200 €,
# nettoyage en régie 40 €/h — voir degats-locatifs/grille-prix-degats-2026.md)
TARIFS_CLES = {"copies": 15, "cyl_std": 135, "cyl_secu": 340, "detecteurs": 30}
TARIFS_ENTRETIEN = {"boiler": 150, "ramonage": 100, "chauffage": 250, "chauffage_gaz": 200, "adoucisseur": 120,
                    "alarme": 180, "ventilation": 150, "fosse": 120}
PU_NETTOYAGE = 40
LIGNE_CHAUFFAGE = 54
NETTOYAGE_ROWS = list(range(61, 67))       # 6 lignes
MAIN_OEUVRE_ROW = 68
FOURNITURES_ROWS = list(range(69, 73))     # 4 lignes
AUTRES_ROWS = list(range(74, 131))         # 57 lignes
SIGNATURE_ANCHOR = "C135"                  # au-dessus de « po L'expert en charge » (C139)

# Lignes qui alimentent la TVA 21 % (Autres, Nettoyage, Extérieurs, expert…)
ROWS_TVA21_FIXES = (
    list(CLES_ROWS.values()) + list(ENTRETIEN_ROWS.values())
    + NETTOYAGE_ROWS + [MAIN_OEUVRE_ROW] + FOURNITURES_ROWS
)


class SortieError(Exception):
    pass


# --------------------------------------------------------------------------
# Utilitaires
# --------------------------------------------------------------------------
_OPS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.USub: operator.neg, ast.UAdd: operator.pos,
}


def eval_montant(expr):
    """Évalue un prix saisi : nombre (« 150 », « 12,5 ») ou petit calcul
    (« (80*15)+3*70 »), avec ou sans « = » devant. Retourne un float, ou
    None si vide. Lève ValueError si illisible."""
    if expr is None:
        return None
    if isinstance(expr, (int, float)):
        return float(expr)
    s = str(expr).strip().replace("€", "").replace(" ", "")
    if not s:
        return None
    if s.startswith("="):
        s = s[1:]
    # virgule décimale belge → point (« 1.250,50 » → « 1250.50 »)
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    node = ast.parse(s, mode="eval").body

    def ev(n):
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
            return float(n.value)
        if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.operand))
        raise ValueError(f"Prix illisible : {expr}")

    return ev(node)


def excel_prix(expr):
    """Valeur à écrire dans la colonne PU : un nombre, ou une formule Excel
    si Georges a saisi un calcul (comme « =(80*15)+3*70 » dans le 2246)."""
    if expr is None or str(expr).strip() == "":
        return None
    s = str(expr).strip().replace("€", "").strip()
    if s.startswith("="):
        s = s[1:]
    if re.fullmatch(r"[0-9.,\s]+", s):
        return eval_montant(s)
    eval_montant(s)  # vérifie que c'est lisible
    return "=" + s.replace(",", ".").replace(" ", "")


def to_num(v, default=0.0):
    if v is None or str(v).strip() == "":
        return default
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return default


def parse_date(s):
    s = (s or "").strip()
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def today_str():
    return date.today().strftime("%d/%m/%Y")


def nom_pv_depuis_excel(nom_excel):
    """« M. DU TOIN Jonathan & Mme SINGH Kiara » → « Jonathan DU TOIN &
    Kiara SINGH » (forme du PV). Simple proposition, modifiable."""
    parts = []
    for bloc in re.split(r"\s*&\s*", nom_excel or ""):
        tokens = [t for t in bloc.split() if t.strip(".").lower() not in ("m", "mme", "mr", "melle", "mlle", "mmes", "mm")]
        noms = [t for t in tokens if t.isupper() and any(ch.isalpha() for ch in t)]
        prenoms = [t for t in tokens if t not in noms]
        if noms and prenoms:
            parts.append(" ".join(prenoms + noms))
        else:
            parts.append(" ".join(tokens))
    return " & ".join(p for p in parts if p)


def _avec_unite(valeur, unite):
    v = (valeur or "").strip()
    if not v:
        return ""
    if re.search(r"[A-Za-zµ³]", v):
        return v
    return f"{v} {unite}"


# --------------------------------------------------------------------------
# Modèle de saisie
# --------------------------------------------------------------------------
def empty_data():
    return {
        "version": 1,
        "dossier": "",
        "initiales": "GB",
        "dossier_folder": "",
        "bien": {"rue": "", "cp": "", "commune": ""},
        "locataire": {"nom_excel": "", "nom_pv": "", "rue": "", "cpville": "", "repr": "Les précités"},
        "bailleur": {"nom_excel": "", "nom_pv": "", "rue": "", "cpville": "", "repr": ""},
        "date_entree": "",
        "date_sortie": "",
        "date_cloture": today_str(),
        "compteurs": {
            "eau": {"num": "", "index1": "", "index2": ""},
            "elec": {"num": "", "index1": "", "index2": ""},
            "gaz": {"num": "", "index1": "", "index2": ""},
            "chaleur": {"num": "", "index1": "", "index2": ""},
        },
        "citerne": {"litres": "", "prix": ""},
        "cles": {k: "" for k in CLES_ROWS},
        "entretiens": {k: "" for k in ENTRETIEN_ROWS},
        "nettoyage": [],          # {texte, qte, unite, pu}
        "main_oeuvre": {"qte": "", "pu": "50"},
        "fournitures": [],        # {texte, qte, pu, usure}
        "autres": [],             # {type: "titre"/"poste", texte, qte, nature, pu, usure, tva}
        "autre_accord": "",
        "indisponibilite": "",
        "coordination": False,
    }


def merge_defaults(data):
    base = empty_data()
    for k, v in (data or {}).items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            for kk, vv in v.items():
                if isinstance(vv, dict) and isinstance(base[k].get(kk), dict):
                    base[k][kk].update(vv)
                else:
                    base[k][kk] = vv
        else:
            base[k] = v
    return base


def save_saisie(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_saisie(path):
    with open(path, "r", encoding="utf-8") as f:
        return merge_defaults(json.load(f))


# --------------------------------------------------------------------------
# Calcul du total (identique aux formules du tableau)
# --------------------------------------------------------------------------
def compute(data):
    """Retourne un dict détaillant le décompte : lignes, sous-totaux, TVA,
    total. Reprend les formules du gabarit (I = QTE × PU × (1 − usure))."""
    lignes = []  # (libellé, total_ligne, taux_tva)

    def add(label, qte, pu, usure=0.0, tva=21):
        q = to_num(qte)
        p = eval_montant(pu) if str(pu or "").strip() else 0.0
        u = to_num(usure) / 100.0 if usure not in (None, "") else 0.0
        tot = round(q * (p or 0.0) * (1 - u), 2)
        lignes.append((label, tot, tva))

    for k, row in CLES_ROWS.items():
        add(CLES_LABELS[k], data["cles"].get(k), TARIFS_CLES[k])
    for k in ENTRETIEN_ROWS:
        add(ENTRETIEN_LABELS[k], data["entretiens"].get(k), TARIFS_ENTRETIEN[k])
    for n in data["nettoyage"]:
        add("Nettoyage : " + (n.get("texte") or ""), n.get("qte"), n.get("pu") or PU_NETTOYAGE)
    add("Main d'œuvre", data["main_oeuvre"].get("qte"), data["main_oeuvre"].get("pu") or 50)
    for f in data["fournitures"]:
        add("Fourniture : " + (f.get("texte") or ""), f.get("qte"), f.get("pu"), f.get("usure"))
    for a in data["autres"]:
        if a.get("type") == "titre":
            continue
        add(a.get("texte") or "", a.get("qte"), a.get("pu"), a.get("usure"), int(to_num(a.get("tva"), 6)))

    total_htva = round(sum(t for _l, t, _r in lignes), 2)
    base6 = round(sum(t for _l, t, r in lignes if r == 6), 2)
    base21 = round(sum(t for _l, t, r in lignes if r == 21), 2)
    tva6 = round(base6 * 0.06, 2)
    tva21 = round(base21 * 0.21, 2)
    autre = round(to_num(data.get("autre_accord")), 2)
    coord = round(total_htva * 0.10, 2) if data.get("coordination") else 0.0
    total = round(total_htva + autre + tva6 + tva21 + coord, 2)
    return {
        "lignes": lignes, "total_htva": total_htva, "autre": autre,
        "coordination": coord, "tva6": tva6, "tva21": tva21, "total": total,
    }


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def gabarit_excel_path():
    p = os.path.join(GABARIT_DIR, "ELLS_gabarit_valorisation.xlsx")
    if os.path.exists(p):
        return p
    for f in sorted(os.listdir(GABARIT_DIR)):
        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
            return os.path.join(GABARIT_DIR, f)
    raise SortieError("Gabarit Excel de sortie introuvable dans le dossier « gabarit_sortie ».")


def gabarit_pv_path():
    p = os.path.join(GABARIT_DIR, "PV_sortie_gabarit.docx")
    if not os.path.exists(p):
        raise SortieError("Gabarit PV Word introuvable dans le dossier « gabarit_sortie ».")
    return p


def _set_text(ws, addr, value):
    """Texte forcé (équivalent de l'apostrophe devant la saisie)."""
    c = ws[addr]
    c.value = value
    if isinstance(value, str) and value:
        c.quotePrefix = True


def fill_excel(data, output_path, template_path=None):
    template_path = template_path or gabarit_excel_path()
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]

    b = data["bien"]
    _set_text(ws, "B6", f"{b['cp']} {b['commune']}, {b['rue']}".strip(" ,"))

    loc, bai = data["locataire"], data["bailleur"]
    _set_text(ws, "B9", loc["nom_excel"])
    _set_text(ws, "B10", loc["rue"])
    _set_text(ws, "B11", loc["cpville"])
    _set_text(ws, "B14", loc["repr"])
    _set_text(ws, "D9", bai["nom_excel"])
    _set_text(ws, "D10", bai["rue"])
    _set_text(ws, "D11", bai["cpville"])
    _set_text(ws, "D14", bai["repr"])

    for addr, s in (("I30", data["date_entree"]), ("I31", data["date_sortie"])):
        d = parse_date(s)
        if d:
            ws[addr].value = d
            ws[addr].number_format = "dd/mm/yyyy"
        elif s:
            _set_text(ws, addr, s)

    # Compteurs
    cpt = data["compteurs"]
    for row, key, label, unite in (
        (34, "eau", "Compteur Eau n° ", "m³"),
        (35, "elec", "Compteur Electricité n° ", "KWh"),
        (36, "gaz", "Compteur Gaz n° ", "m³"),
    ):
        c = cpt[key]
        if c["num"]:
            _set_text(ws, f"B{row}", label + c["num"])
        if c["index1"]:
            _set_text(ws, f"C{row}", _avec_unite(c["index1"], unite))
        if c["index2"]:
            _set_text(ws, f"G{row}", _avec_unite(c["index2"], unite))
    ch = cpt["chaleur"]
    if ch["num"]:
        _set_text(ws, "B40", "Compteur Chaleur : S/N: " + ch["num"])
    if ch["index1"]:
        _set_text(ws, "C40", _avec_unite(ch["index1"], "KWh"))
    if ch["index2"]:
        _set_text(ws, "G40", _avec_unite(ch["index2"], "KWh"))

    # Citerne (décompte entre parties, hors total comme dans le gabarit)
    cit = data["citerne"]
    if str(cit.get("litres") or "").strip():
        ws["C38"].value = to_num(cit["litres"])
        if str(cit.get("prix") or "").strip():
            ws["H38"].value = to_num(cit["prix"])
        ws["I38"].value = "=C38*H38"

    # Clés / entretiens : seule la quantité est reportée (tarifs du gabarit)
    for k, row in CLES_ROWS.items():
        q = data["cles"].get(k)
        if str(q or "").strip():
            ws[f"C{row}"].value = to_num(q)
    for k, row in ENTRETIEN_ROWS.items():
        if row == LIGNE_CHAUFFAGE:
            continue
        q = data["entretiens"].get(k)
        if str(q or "").strip():
            ws[f"C{row}"].value = to_num(q)
    # Une seule ligne « Entretien chauffage » dans le tableau : mazout (250 €)
    # ou gaz (200 €), libellé et prix adaptés.
    q_maz = to_num(data["entretiens"].get("chauffage"))
    q_gaz = to_num(data["entretiens"].get("chauffage_gaz"))
    if q_maz and q_gaz:
        raise SortieError("Entretien chaudière : mazout OU gaz (une seule ligne dans le tableau).")
    if q_maz or q_gaz:
        k = "chauffage_gaz" if q_gaz else "chauffage"
        r = LIGNE_CHAUFFAGE
        _set_text(ws, f"B{r}", ENTRETIEN_LABELS[k])
        ws[f"C{r}"].value = q_gaz or q_maz
        ws[f"G{r}"].value = TARIFS_ENTRETIEN[k]

    # Nettoyage
    if len(data["nettoyage"]) > len(NETTOYAGE_ROWS):
        raise SortieError(f"Maximum {len(NETTOYAGE_ROWS)} lignes de nettoyage.")
    for row, n in zip(NETTOYAGE_ROWS, data["nettoyage"]):
        if n.get("texte"):
            _set_text(ws, f"B{row}", n["texte"])
        ws[f"C{row}"].value = to_num(n.get("qte"))
        _set_text(ws, f"D{row}", n.get("unite") or "forfait")
        pu = excel_prix(n.get("pu"))
        ws[f"G{row}"].value = pu if pu is not None else PU_NETTOYAGE
        ws[f"I{row}"].value = f"=G{row}*C{row}"

    mo = data["main_oeuvre"]
    if str(mo.get("qte") or "").strip():
        ws[f"C{MAIN_OEUVRE_ROW}"].value = to_num(mo["qte"])
        pu = excel_prix(mo.get("pu"))
        if pu is not None:
            ws[f"G{MAIN_OEUVRE_ROW}"].value = pu

    # Petites fournitures
    if len(data["fournitures"]) > len(FOURNITURES_ROWS):
        raise SortieError(f"Maximum {len(FOURNITURES_ROWS)} lignes de petites fournitures.")
    for i, (row, f) in enumerate(zip(FOURNITURES_ROWS, data["fournitures"])):
        texte = f.get("texte") or ""
        if i == 0 and row == 69:
            texte = "Petites fournitures : " + texte if texte else "Petites fournitures :"
        _set_text(ws, f"B{row}", texte)
        if str(f.get("qte") or "").strip():
            ws[f"C{row}"].value = to_num(f["qte"])
        pu = excel_prix(f.get("pu"))
        ws[f"G{row}"].value = pu if pu is not None else None
        if str(f.get("usure") or "").strip():
            ws[f"H{row}"].value = to_num(f["usure"]) / 100.0
            ws[f"H{row}"].number_format = "0%"
            ws[f"I{row}"].value = f"=C{row}*G{row}*(1-H{row})"
        else:
            ws[f"I{row}"].value = f"=G{row}*C{row}"

    # Autres postes
    if len(data["autres"]) > len(AUTRES_ROWS):
        raise SortieError(f"Maximum {len(AUTRES_ROWS)} lignes dans « Autres postes ».")
    rows6, rows21 = [], []
    for row, a in zip(AUTRES_ROWS, data["autres"]):
        texte = a.get("texte") or ""
        if a.get("type") == "titre":
            _set_text(ws, f"B{row}", texte)
            for col in "CGHI":
                ws[f"{col}{row}"].value = None
            continue
        if texte and not texte.startswith("-"):
            texte = "- " + texte
        _set_text(ws, f"B{row}", texte)
        ws[f"C{row}"].value = to_num(a.get("qte"))
        if a.get("nature"):
            _set_text(ws, f"D{row}", a["nature"])
        pu = excel_prix(a.get("pu"))
        ws[f"G{row}"].value = pu  # None = prix non encore fixé
        ws[f"H{row}"].value = to_num(a.get("usure")) / 100.0
        ws[f"I{row}"].value = f"=C{row}*G{row}*(1-H{row})"
        taux = int(to_num(a.get("tva"), 6))
        (rows6 if taux == 6 else rows21 if taux == 21 else []).append(row)

    # Récapitulatif
    if str(data.get("autre_accord") or "").strip():
        ws["I21"].value = to_num(data["autre_accord"])
    if data.get("indisponibilite"):
        _set_text(ws, "D22", data["indisponibilite"])
    if data.get("coordination"):
        ws["I23"].value = "=I20*0.1"
    rows21_all = ROWS_TVA21_FIXES + rows21
    ws["I24"].value = ("=(" + "+".join(f"I{r}" for r in rows6) + ")*0.06") if rows6 else 0
    ws["I25"].value = "=(" + "+".join(f"I{r}" for r in rows21_all) + ")*0.21"

    # Signature (autorisée par Georges) au-dessus de « po L'expert en charge »
    sig = signature.signature_path()
    if sig:
        try:
            from openpyxl.drawing.image import Image as XLImage

            img = XLImage(sig)
            img.width, img.height = 170, 85
            ws.add_image(img, SIGNATURE_ANCHOR)
        except Exception:  # noqa: BLE001 — la signature n'est jamais bloquante
            pass

    wb.save(output_path)
    return output_path


# --------------------------------------------------------------------------
# PV Word
# --------------------------------------------------------------------------
def _repr_ligne(repr_txt):
    """« Les précités » (Excel) → « Représenté par le précité » (comme dans
    le PV 2246) ; un nom → « Représenté par : <nom> »."""
    r = (repr_txt or "").strip()
    if not r or "précit" in r.lower():
        return "Représenté par le précité"
    return f"Représenté par : {r}"


def fill_pv(data, output_path, total, template_path=None):
    template_path = template_path or gabarit_pv_path()
    loc, bai, b = data["locataire"], data["bailleur"], data["bien"]
    scalars = {
        "Loc_nom": loc["nom_pv"] or nom_pv_depuis_excel(loc["nom_excel"]),
        "Loc_rue": loc["rue"],
        "Loc_cpville": loc["cpville"],
        "Loc_repr": _repr_ligne(loc["repr"]),
        "Bai_nom": bai["nom_pv"] or nom_pv_depuis_excel(bai["nom_excel"]),
        "Bai_rue": bai["rue"],
        "Bai_cpville": bai["cpville"],
        "Bai_repr": _repr_ligne(bai["repr"]),
        "Date_sortie": data["date_sortie"],
        "Date_entree": data["date_entree"],
        "Bien_rue": b["rue"],
        "Bien_cp": b["cp"],
        "Bien_commune": b["commune"],
        "Montant_chiffres": montant_lettres.montant_en_chiffres(total),
        "Montant_lettres": montant_lettres.montant_en_lettres(total),
        "Date_cloture": data.get("date_cloture") or today_str(),
        "Dossier": data["dossier"],
    }
    unknown = docx_fill.fill_docx(template_path, output_path, scalars, {})
    _add_pv_signature(output_path)
    return unknown


def _add_pv_signature(pv_path):
    """Signature (autorisée) au-dessus de « Po l'expert en charge »."""
    sig = signature.signature_path()
    if not sig:
        return
    import docx
    from docx.shared import Emu

    d = docx.Document(pv_path)
    paras = d.paragraphs
    for i, p in enumerate(paras):
        if "expert en charge" in p.text.lower() and i > 0:
            target = paras[i - 1]
            if not target.text.strip():
                target.add_run().add_picture(sig, width=Emu(1619250), height=Emu(809625))
                d.save(pv_path)
            return


# --------------------------------------------------------------------------
# Noms de fichiers et dossiers
# --------------------------------------------------------------------------
_DATE_SUFFIX_RE = re.compile(r"\s*\(\s*\d{1,2}[.\-/]\d{1,2}\s*\)\s*")


def parse_folder_name(folder):
    """« 2026.2766 ELLS Hannut Rue de Tirlemont 53-2 (01.10) GB » →
    {dossier, type, reste, initiales, base}."""
    name = os.path.basename(os.path.normpath(folder or ""))
    m = re.match(r"^(\d{4}\.\d{3,4})\s+(\S+)\s+(.*)$", name)
    if not m:
        return {"dossier": "", "type": "", "reste": "", "initiales": "", "base": name}
    dossier, typ, reste = m.groups()
    initiales = ""
    m2 = re.search(r"(?:^|[\s)])\(?([A-Za-z]{2,4})\)?\s*$", reste)
    if m2:
        initiales = m2.group(1).upper()
        reste = reste[: m2.start(1)].rstrip(" (")
    reste = _DATE_SUFFIX_RE.sub(" ", reste).strip()
    return {"dossier": dossier, "type": typ, "reste": reste, "initiales": initiales,
            "base": f"{dossier} {typ} {reste}".strip()}


def _safe(name):
    return re.sub(r'[\\/:*?"<>|]+', "-", name).strip()


def output_names(data):
    """(nom Excel, nom PV Word), sur le modèle du 2246 :
    « 2026.2246 ELLS Dégâts locs Woluwé-Saint-Lambert Avenue de Broqueville 235-6 GB.xlsx »
    « 2026.2246 ELLS WSL Av de Broqueville 235-6 GB.docx »"""
    ini = data.get("initiales") or "GB"
    b = data["bien"]
    rue = b["rue"].replace("/", "-")
    xlsx = _safe(f"{data['dossier']} ELLS Dégâts locs {b['commune']} {rue} {ini}") + ".xlsx"
    info = parse_folder_name(data.get("dossier_folder"))
    if info["reste"]:
        pv = _safe(f"{data['dossier']} ELLS {info['reste']} {ini}") + ".docx"
    else:
        pv = _safe(f"{data['dossier']} ELLS {b['commune']} {rue} {ini}") + ".docx"
    return xlsx, pv


def existing_outputs(data, output_dir=None):
    out_dir = output_dir or data.get("dossier_folder")
    return [p for p in (os.path.join(out_dir or "", n) for n in output_names(data)) if os.path.exists(p)]


def generate(data, output_dir=None, overwrite=False):
    """Génère l'Excel et le PV Word dans le dossier du n° (ou output_dir).
    N'écrase un fichier existant que si overwrite=True, et toujours après
    une sauvegarde hors Dropbox. Retourne {"excel", "pv", "total", "calc",
    "saisie", "sauvegardes"}."""
    if not data.get("dossier"):
        raise SortieError("Le n° de dossier est obligatoire.")
    out_dir = output_dir or data.get("dossier_folder")
    if not out_dir or not os.path.isdir(out_dir):
        raise SortieError("Choisissez le dossier du n° (où déposer les fichiers).")
    securite.check_dossier_folder(out_dir)
    calc = compute(data)
    xlsx_name, pv_name = output_names(data)
    xlsx_path = os.path.join(out_dir, xlsx_name)
    pv_path = os.path.join(out_dir, pv_name)
    saisie_path = os.path.join(out_dir, f"{data['dossier']} saisie sortie.json")
    exist = [p for p in (xlsx_path, pv_path) if os.path.exists(p)]
    if exist and not overwrite:
        raise SortieError("Ces fichiers existent déjà : " + ", ".join(os.path.basename(p) for p in exist))
    for p in (xlsx_path, pv_path, saisie_path):
        securite.ensure_not_locked(p)
    sauvegardes = [securite.backup_file(p) for p in (xlsx_path, pv_path, saisie_path) if os.path.exists(p)]

    tmp_x, tmp_pv = securite.temp_path_for(xlsx_path), securite.temp_path_for(pv_path)
    try:
        fill_excel(data, tmp_x)
        fill_pv(data, tmp_pv, calc["total"])
        openpyxl.load_workbook(tmp_x)  # vérifie que les fichiers se relisent
        import docx as _docx_check

        _docx_check.Document(tmp_pv)
    except Exception:
        for t in (tmp_x, tmp_pv):
            if os.path.exists(t):
                os.remove(t)  # fichiers temporaires du programme uniquement
        raise
    securite.commit_temp(tmp_x, xlsx_path)
    securite.commit_temp(tmp_pv, pv_path)
    save_saisie(data, saisie_path)
    return {"excel": xlsx_path, "pv": pv_path, "total": calc["total"], "calc": calc,
            "saisie": saisie_path, "sauvegardes": [s for s in sauvegardes if s]}
