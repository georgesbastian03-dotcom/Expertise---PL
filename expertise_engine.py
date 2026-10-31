"""Expertise immobilière (EXP) — 100 % local, sans IA.

L'expert garde la main sur la lecture des notes, l'analyse et les valeurs.
Ce module fait la partie mécanique de l'encodage :

  * ABEX : copie du modèle VIERGE officiel (« 2026.0000 ABEX Calculs 1056
    2026 Officiel.xlsx » — jamais l'ABEX d'un autre dossier) et remplissage
    du seul onglet « Donnees » (type, PEB, chambres, sanitaires, n° de
    dossier, date de visite, adresse, prix et année d'achat, RC, année de
    construction, surfaces par niveau, terrain). Le fichier est modifié
    directement dans son XML : formules, images, graphiques et autres
    onglets restent intacts ; Excel recalcule tout à l'ouverture ;
  * photos : façade en page de garde (cadre existant du Word, rotation du
    modèle gérée) et planches 3 × 3 sous « PHOTO(S) DE L'IMMEUBLE » ;
  * jamais d'écrasement d'un fichier sur lequel l'expert a travaillé : une
    nouvelle version (« … v2 ») est écrite à côté.
"""
import io
import json
import os
import re
import zipfile
from copy import deepcopy
from datetime import date, datetime

from PIL import Image, ImageOps

import config
import entree2027 as e27
import entree_engine as ee
import securite

W = e27.W
A = e27.A
R = e27.R
EMU_PER_CM = 360000

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
X = "{%s}" % NS_MAIN

MODELE_ABEX_NOM = "2026.0000 ABEX Calculs 1056 2026 Officiel.xlsx"
FEUILLE = "Donnees"

# Champs de l'onglet Données : clé → (cellule, libellé, type)
CHAMPS = [
    ("type", "B2", "Type de bien", "texte"),
    ("peb_classe", "E2", "Classe PEB", "texte"),
    ("peb_m2", "E3", "m² du PEB (1re page)", "nombre"),
    ("chambres", "E4", "Nombre de chambres", "nombre"),
    ("sanitaires", "E5", "Nombre de sanitaires", "nombre"),
    ("date_visite", "C3", "Date de visite (jj/mm/aaaa)", "date"),
    ("rue", "B6", "Rue et n°", "texte"),
    ("localite", "B7", "CP et localité", "texte"),
    ("prix_achat", "G4", "Prix d'achat", "nombre"),
    ("annee_achat", "G5", "Année d'achat", "nombre"),
    ("rc", "G7", "RC de base", "nombre"),
    ("annee_construction", "G8", "Année de construction", "nombre"),
    ("terrain_m2", "G42", "Superficie du terrain (m²) — valeur écrite", "nombre"),
]
# Référence cadastrale (WalOnMap, encadré « Référence cadastrale », ou
# CADGIS). Règle GB 31/10 : le cadastre officiel l'emporte sur les notes,
# l'OM et le client pour IDENTIFIER la parcelle ; CADGIS l'emporte sur
# CADMAP en cas de divergence. La superficie est toujours une valeur ÉCRITE
# et officielle (CADGIS, plan, notes), jamais une surface calculée sur le
# contour CADMAP ; si plusieurs documents diffèrent, on retient la valeur qui
# revient le plus souvent (exemple 3490 : 8 a 23, et non 818 m² calculés).
CADASTRE = [
    ("capakey", "CaPaKey"),
    ("division", "Division"),
    ("section_cad", "Section"),
    ("radical", "Radical"),
    ("exposant", "Exposant"),
    ("puissance", "Puissance"),
    ("source_cad", "Source de la parcelle (CADGIS / WalOnMap / CADMAP)"),
    ("source_sup", "Source de la superficie (CADGIS / plan / notes)"),
]
# Ce que disent les notes, l'OM ou le client : seulement pour comparer.
NOTES_CAD = [
    ("parcelle_notes", "Parcelle selon les notes / l'OM (ex. A 53)"),
    ("terrain_notes", "Superficie selon un autre document (m²)"),
]
# Surfaces (SEM) par niveau : lignes C11 à C32 du modèle
SURFACES = [
    ("C11", "Caves non aménagées"), ("C12", "Garage sous-sol"), ("C14", "Locaux sous-sol aménagés"),
    ("C16", "Annexe non habitable"), ("C17", "Annexe habitable"), ("C19", "Garage rez"),
    ("C21", "Rez principal"), ("C22", "Bel étage"), ("C23", "Véranda hab. / extension"),
    ("C24", "Premier étage (> 2,10)"), ("C25", "Autres étages (> 2,10)"), ("C26", "Étage mansardé 1,5 < 2,1"),
    ("C27", "Réserve aménageable"), ("C28", "Combles / grenier utile"), ("C29", "Toit plat / carport aménagé"),
    ("C30", "Surface terrasse"), ("C31", "Parking intérieur"), ("C32", "Parking extérieur"),
]
# Valeurs d'exemple du modèle vierge à toujours écraser (C21 = 1, G42 = 800)
A_ECRASER = {"C21": 0, "G42": 0}

PLANCHE_PX = (2400, 1800)
PLANCHE_W_CM = 16.0
PLANCHE_GAP_PX = 14


class ExpertiseError(Exception):
    pass


# --------------------------------------------------------------------------
# Dossier
# --------------------------------------------------------------------------
def parse_folder(folder):
    return ee.parse_folder_name(folder)


def output_names(folder, initiales="GB"):
    info = parse_folder(folder)
    ini = info["initiales"] or initiales
    base = re.sub(r'[\\/:*?"<>|]+', "-", f"{info['reste']} {ini}").strip()
    return {"abex": f"{info['dossier']} EXP ABEX {base}.xlsx", "word": f"{info['dossier']} EXP {base}.docx"}


def next_version(path):
    """« X.docx » libre → X.docx ; sinon « X v2.docx », « X v3.docx »…"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    base = re.sub(r"\s+v\d+$", "", base)
    k = 2
    while os.path.exists(f"{base} v{k}{ext}"):
        k += 1
    return f"{base} v{k}{ext}"


def find_words(folder):
    return [os.path.join(folder, f) for f in sorted(os.listdir(folder))
            if f.lower().endswith(".docx") and not f.startswith("~$") and " EXP " in f.upper()]


def find_abex_model(cfg):
    p = cfg.get("modele_abex") or ""
    return p if p and os.path.exists(p) else None


def default_data(folder):
    info = parse_folder(folder)
    m = re.match(r"^(\S+(?:-\S+)*)\s+(.*)$", info["reste"])
    return {"dossier": info["dossier"], "rue": m.group(2) if m else info["reste"],
            "localite": m.group(1) if m else "", "surfaces": {}}


def _saisie_path(dossier):
    d = config._config_dir() / "saisies_expertise"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"{dossier or 'sans_numero'}.json"


def save_saisie(data):
    _saisie_path(data.get("dossier")).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_saisie(dossier):
    p = _saisie_path(dossier)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            return None
    return None


# --------------------------------------------------------------------------
# Valeurs saisies
# --------------------------------------------------------------------------
def parse_number(txt):
    """« 1 250,5 » → 1250.5 ; « =9*6,98+2 » → formule « 9*6.98+2 » ;
    vide → None."""
    s = (txt or "").strip()
    if not s:
        return None
    if s.startswith("="):
        return ("formule", s[1:].replace(",", ".").replace(" ", ""))
    s2 = s.replace(" ", "").replace(" ", "").replace("€", "").replace(",", ".")
    try:
        v = float(s2)
        return int(v) if v.is_integer() else v
    except ValueError:
        raise ExpertiseError(f"« {txt} » n'est pas un nombre (formule : commencer par =).")


def excel_date(txt):
    m = re.match(r"\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\s*$", txt or "")
    if not m:
        raise ExpertiseError(f"Date « {txt} » non reconnue (jj/mm/aaaa).")
    j, mo, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
    a += 2000 if a < 100 else 0
    return (date(a, mo, j) - date(1899, 12, 30)).days


def cell_values(data):
    """{cellule: valeur} à écrire dans l'onglet Données. valeur = str
    (texte), nombre, ("formule", expr) ou None (cellule vidée)."""
    out = {}
    out["G2"] = (data.get("dossier") or "").replace(".", "-")
    for key, cell, label, kind in CHAMPS:
        txt = (data.get(key) or "").strip() if isinstance(data.get(key), str) else data.get(key)
        if kind == "texte":
            out[cell] = txt or None
        elif kind == "date":
            out[cell] = excel_date(txt) if txt else None
        else:
            try:
                out[cell] = parse_number(txt if isinstance(txt, str) else ("" if txt is None else str(txt)))
            except ExpertiseError as exc:
                raise ExpertiseError(f"{label} : {exc}") from exc
    m = re.match(r"\s*(\d{4})", data.get("localite") or "")
    if m:
        out["G41"] = int(m.group(1))
    for cell, label in SURFACES:
        txt = (data.get("surfaces") or {}).get(cell, "")
        try:
            v = parse_number(txt)
        except ExpertiseError as exc:
            raise ExpertiseError(f"Surface « {label} » : {exc}") from exc
        out[cell] = v
    for cell, v in A_ECRASER.items():
        if out.get(cell) is None:
            out[cell] = v
    return out


# --------------------------------------------------------------------------
# Cadastre
# --------------------------------------------------------------------------
_CAPAKEY_RE = re.compile(r"(\d{5})\s*([A-Z])\s*(\d{4})\s*/\s*(\d{2})\s*([A-Z_]?)\s*(\d{3})")


def parse_capakey(txt):
    """« 25005A0051/00S000 » → {division, section_cad, radical, exposant,
    puissance} ; None si le texte n'est pas une CaPaKey."""
    m = _CAPAKEY_RE.search((txt or "").upper())
    if not m:
        return None
    div, sec, rad, bis, exp, pui = m.groups()
    out = {"capakey": f"{div}{sec}{rad}/{bis}{exp}{pui}", "division": div, "section_cad": sec,
           "radical": rad, "exposant": exp.replace("_", ""), "puissance": pui}
    if int(bis):
        out["bis"] = bis
    return out


def parse_superficie(txt):
    """« 1272 m² », « 1.272 m² », « 12 a 72 ca », « 0 ha 12 a 72 ca » → m²
    (float) ; None si rien de reconnu."""
    t = (txt or "").lower().replace("\u00a0", " ")
    m = re.search(r"(?:(\d+)\s*ha)?\s*(?:(\d+)\s*a(?![a-z]))?\s*(\d+)\s*ca\b", t)
    if m and (m.group(1) or m.group(2) or m.group(3)):
        return int(m.group(1) or 0) * 10000 + int(m.group(2) or 0) * 100 + int(m.group(3) or 0)
    m = re.search(r"(\d[\d .]*(?:,\d+)?)\s*m(?:²|2)", t)
    if m:
        return float(m.group(1).replace(" ", "").replace(".", "").replace(",", "."))
    return None


def parse_walonmap(txt):
    """Texte copié de l'encadré « Référence cadastrale » de WalOnMap (ou de
    CADGIS) → champs reconnus (CaPaKey et ses parties, superficie)."""
    out = parse_capakey(txt) or {}
    if not out:
        for key, label in (("division", "division"), ("section_cad", "section"), ("radical", "radical"),
                           ("exposant", "exposant"), ("puissance", "puissance")):
            m = re.search(label + r"\s*[:=]?\s*([0-9A-Za-z_]+)", txt or "", re.I)
            if m:
                out[key] = m.group(1).upper()
    m = re.search(r"(?:superficie|surface)[^\n:]*[:=]?\s*([^\n]+)", txt or "", re.I)
    sup = parse_superficie(m.group(1) if m else txt)
    if sup is not None and re.search(r"calcul", (m.group(0) if m else txt) or "", re.I):
        out["sup_calculee"] = True  # surface calculée sur le contour : jamais reprise
    elif sup is not None:
        out["terrain_m2"] = ("%g" % sup).replace(".", ",")
    return out


def parcelle_label(data):
    """Libellé court de la parcelle : « A 0051S » (section, radical, exposant)."""
    sec = (data.get("section_cad") or "").strip().upper()
    rad = (data.get("radical") or "").strip()
    exp = (data.get("exposant") or "").strip().upper()
    if rad.isdigit():
        rad = rad.zfill(4)
    return " ".join(x for x in (sec, rad + exp) if x)


def _norm_parcelle(txt):
    """« A 53 », « A0053 », « 53 » → ("A", 53, "") pour comparer."""
    t = (txt or "").upper().replace(" ", "")
    m = re.match(r"^([A-Z]?)0*(\d+)([A-Z]?)", t)
    return (m.group(1), int(m.group(2)), m.group(3)) if m else None


def ecarts_cadastre(data):
    """Écarts entre le cadastre officiel et les notes / l'OM, à signaler à
    GB (pour la parcelle, le cadastre est retenu sans attendre sa décision)."""
    out = []
    lab = parcelle_label(data)
    notes = (data.get("parcelle_notes") or "").strip()
    if notes and lab:
        a, b = _norm_parcelle(notes), _norm_parcelle(lab)
        if a and b and (a[1] != b[1] or (a[0] and a[0] != b[0]) or (a[2] and a[2] != b[2])):
            out.append(f"Parcelle : les notes indiquent « {notes} », le cadastre « {lab} » (retenu).")
    try:
        cad = parse_number(data.get("terrain_m2") or "")
        nt = parse_number(data.get("terrain_notes") or "")
    except ExpertiseError:
        cad = nt = None
    if isinstance(cad, (int, float)) and isinstance(nt, (int, float)) and abs(cad - nt) > 0.5:
        out.append(f"Superficie : G42 = {cad:g} m², un autre document indique {nt:g} m² — retenez la valeur "
                   "écrite qui revient le plus souvent (CADGIS, plan, notes).")
    if "CADMAP" in (data.get("source_sup") or "").upper():
        out.append("Superficie : jamais une surface calculée sur le contour CADMAP — reprenez une valeur écrite "
                   "(CADGIS, plan ou notes).")
    if notes and not lab:
        out.append("Parcelle des notes encodée, mais pas la référence cadastrale officielle (WalOnMap / CADGIS).")
    return out


# --------------------------------------------------------------------------
# Écriture XLSX (XML direct : rien d'autre n'est touché)
# --------------------------------------------------------------------------
def _col_row(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return col, int(m.group(2))


def _sheet_path(z, name):
    from lxml import etree

    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for s in wb.iter(X + "sheet"):
        if s.get("name") == name:
            rid = s.get("{%s}id" % NS_REL)
    if rid is None:
        raise ExpertiseError(f"Onglet « {name} » introuvable dans l'ABEX.")
    for r in rels.iter("{%s}Relationship" % NS_PKG_REL):
        if r.get("Id") == rid:
            t = r.get("Target")
            return t.lstrip("/") if t.startswith("/") else "xl/" + t
    raise ExpertiseError("Structure de l'ABEX inattendue.")


def _set_cell(sheet_data, ref, value):
    from lxml import etree

    col, row = _col_row(ref)
    row_el = None
    for r in sheet_data.findall(X + "row"):
        n = int(r.get("r"))
        if n == row:
            row_el = r
            break
        if n > row:
            row_el = etree.Element(X + "row", r=str(row))
            r.addprevious(row_el)
            break
    if row_el is None:
        row_el = etree.SubElement(sheet_data, X + "row", r=str(row))
    cell = None
    for c in row_el.findall(X + "c"):
        cc, _ = _col_row(c.get("r"))
        if cc == col:
            cell = c
            break
        if cc > col:
            cell = etree.Element(X + "c", r=ref)
            c.addprevious(cell)
            break
    if cell is None:
        cell = etree.SubElement(row_el, X + "c", r=ref)
    style = cell.get("s")
    for ch in list(cell):
        cell.remove(ch)
    for att in list(cell.attrib):
        if att not in ("r", "s"):
            del cell.attrib[att]
    if style is not None:
        cell.set("s", style)
    if value is None:
        return
    if isinstance(value, tuple) and value[0] == "formule":
        etree.SubElement(cell, X + "f").text = value[1]
    elif isinstance(value, (int, float)):
        etree.SubElement(cell, X + "v").text = repr(value) if isinstance(value, float) else str(value)
    else:
        cell.set("t", "inlineStr")
        is_ = etree.SubElement(cell, X + "is")
        t = etree.SubElement(is_, X + "t")
        t.text = str(value)
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def fill_abex(model_path, output_path, values):
    """Copie le modèle vierge vers `output_path` en remplissant l'onglet
    Données. Excel recalcule les formules à l'ouverture."""
    from lxml import etree

    with zipfile.ZipFile(model_path) as zin:
        sheet = _sheet_path(zin, FEUILLE)
        root = etree.fromstring(zin.read(sheet))
        sd = root.find(X + "sheetData")
        for ref, v in values.items():
            _set_cell(sd, ref, v)
        new_sheet = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        wb = etree.fromstring(zin.read("xl/workbook.xml"))
        calc = wb.find(X + "calcPr")
        if calc is None:
            calc = etree.SubElement(wb, X + "calcPr")
        calc.set("fullCalcOnLoad", "1")
        new_wb = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)
        temp = securite.temp_path_for(output_path)
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == sheet:
                    zout.writestr(item, new_sheet)
                elif item.filename == "xl/workbook.xml":
                    zout.writestr(item, new_wb)
                elif item.filename == "xl/calcChain.xml":
                    continue  # recalculée par Excel
                else:
                    zout.writestr(item, zin.read(item.filename))
    # calcChain retiré : on enlève aussi sa déclaration
    _drop_calc_chain(temp)
    securite.commit_temp(temp, output_path)
    return output_path


def _drop_calc_chain(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        ct = z.read("[Content_Types].xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    if "calcChain" not in ct and "calcChain" not in rels:
        return
    ct2 = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", ct)
    rels2 = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", rels)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            data = zin.read(n)
            if n == "[Content_Types].xml":
                data = ct2.encode("utf-8")
            elif n == "xl/_rels/workbook.xml.rels":
                data = rels2.encode("utf-8")
            zout.writestr(zin.getinfo(n), data)
    os.replace(tmp, path)


def create_abex(folder, data, model_path, initiales="GB"):
    """Crée l'ABEX du dossier depuis le modèle vierge. Ne remplace jamais un
    ABEX existant : une nouvelle version est écrite à côté. Retourne
    {"path", "remarques"}."""
    securite.check_dossier_folder(folder)
    if not model_path or not os.path.exists(model_path):
        raise ExpertiseError("Modèle ABEX vierge introuvable : choisissez-le (« 2026.0000 ABEX Calculs 1056 2026 "
                             "Officiel.xlsx »).")
    if "0000" not in os.path.basename(model_path):
        raise ExpertiseError("Le modèle choisi ne semble pas être l'ABEX VIERGE (« 2026.0000 … ») : par sécurité, "
                             "on ne part jamais de l'ABEX d'un autre dossier.")
    values = cell_values(data)
    remarques = []
    if values.get("G8") is None:
        remarques.append("Année de construction non remplie : cellule G8 laissée vide (le modèle mettait 1990).")
    if values.get("C3") is None:
        remarques.append("Date de visite non remplie : C3 laissée vide (le modèle mettait la date du jour).")
    remarques += ecarts_cadastre(data)
    if values.get("G42") in (None, 0) and data.get("terrain_notes"):
        remarques.append("Superficie du terrain non encodée en G42 (= 0) alors qu'un autre document en donne une : "
                         "reprenez la valeur écrite qui revient le plus souvent (CADGIS, plan, notes).")
    target = next_version(os.path.join(folder, output_names(folder, initiales)["abex"]))
    if os.path.basename(target) != output_names(folder, initiales)["abex"]:
        remarques.append(f"Un ABEX existait déjà : nouvelle version « {os.path.basename(target)} » (l'existant n'est pas "
                         "touché).")
    fill_abex(model_path, target, values)
    save_saisie(data)
    return {"path": target, "remarques": remarques}


# --------------------------------------------------------------------------
# Photos dans le Word
# --------------------------------------------------------------------------
def _open_photo(path, rot=0):
    im = Image.open(path)
    im = ImageOps.exif_transpose(im).convert("RGB")
    if rot:
        im = im.rotate(-rot, expand=True)
    return im


def make_planche(items):
    """Planche 3 × 3 (au plus 9 photos), cases égales, photos entières
    centrées dans leur case. Retourne les octets JPEG."""
    W_, H_ = PLANCHE_PX
    g = PLANCHE_GAP_PX
    cw, ch = (W_ - 2 * g) / 3, (H_ - 2 * g) / 3
    sheet = Image.new("RGB", PLANCHE_PX, "white")
    for k, it in enumerate(items[:9]):
        im = _open_photo(it["path"], it.get("rot", 0))
        im.thumbnail((int(cw), int(ch)), Image.LANCZOS)
        x0 = int((k % 3) * (cw + g) + (cw - im.width) / 2)
        y0 = int((k // 3) * (ch + g) + (ch - im.height) / 2)
        sheet.paste(im, (x0, y0))
    buf = io.BytesIO()
    sheet.save(buf, "JPEG", quality=85, optimize=True)
    return buf.getvalue()


def _cover_drawing(body):
    """Cadre photo de la page de garde : premier dessin avant « Localisation
    de l'immeuble »."""
    for el in body.iterchildren():
        if "LOCALISATION DE L" in ee._norm(e27._text(el)):
            break
        dr = el.find(".//" + W + "drawing")
        if dr is not None and dr.find(".//" + A + "blip") is not None and el.tag == W + "p":
            return dr
    return None


def _replace_cover(document, drawing, item):
    """Remplace l'image du cadre de garde en gardant le cadre : la photo est
    recadrée au format du cadre affiché et tournée pour compenser la
    rotation du modèle (le modèle Namur affiche l'image tournée de 90°)."""
    ext = drawing.find(".//" + e27.WP + "extent")
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    xfrm = drawing.find(".//" + A + "xfrm")
    rot = int(xfrm.get("rot", "0") or 0) // 60000 if xfrm is not None else 0
    disp_w, disp_h = (cy, cx) if rot % 180 == 90 else (cx, cy)
    im = _open_photo(item["path"], item.get("rot", 0))
    target = disp_w / disp_h
    w, h = im.size
    if w / h > target:
        nw = int(h * target)
        im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    else:
        nh = int(w / target)
        im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    if rot:
        im = im.rotate(rot, expand=True)  # Word la tournera de +rot
    im.thumbnail((1800, 1800), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85, optimize=True)
    blip = drawing.find(".//" + A + "blip")
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    old = blip.get(R + "embed")
    part = Part(PackURI(f"/word/media/couverture_{datetime.now():%H%M%S}.jpeg"), "image/jpeg", buf.getvalue(),
                document.part.package)
    rid = document.part.relate_to(part, RT.IMAGE)
    blip.set(R + "embed", rid)
    return old


def insert_photos(word_path, cover, planche_items, progress_cb=None):
    """Écrit une nouvelle version du Word (« … v2.docx ») avec la photo de
    façade en garde et les planches 3 × 3 sous « PHOTO(S) DE
    L'IMMEUBLE ». Le Word d'origine n'est pas modifié. Retourne {"path",
    "planches", "remarques"}."""
    import docx

    def say(msg):
        if progress_cb:
            progress_cb(msg)

    folder = os.path.dirname(word_path)
    securite.check_dossier_folder(folder)
    d = docx.Document(word_path)
    body = d.element.body
    remarques = []
    removed = set()
    if cover:
        dr = _cover_drawing(body)
        if dr is None:
            remarques.append("Cadre photo de la page de garde introuvable : façade non placée.")
        else:
            say("Photo de façade…")
            removed.add(_replace_cover(d, dr, cover))
    planches = 0
    if planche_items:
        els = list(body.iterchildren())
        i0 = next((i for i, el in enumerate(els) if el.tag == W + "p"
                   and ee._norm(e27._text(el)).startswith("PHOTO(S) DE L")), None)
        if i0 is None:
            raise ExpertiseError("Titre « PHOTO(S) DE L'IMMEUBLE » introuvable dans le Word.")
        # Après le titre : sous-titre éventuel, puis les anciennes planches
        k = i0 + 1
        while k < len(els) and els[k].tag == W + "p" and e27._text(els[k]).strip() and not e27._has_drawing(els[k]):
            k += 1
        anchor = els[k - 1]
        ref_ppr = None
        j = k
        while j < len(els) and els[j].tag == W + "p" and not e27._text(els[j]).strip():
            if e27._has_drawing(els[j]):
                if ref_ppr is None and els[j].find(W + "pPr") is not None:
                    ref_ppr = deepcopy(els[j].find(W + "pPr"))
                e27._remove_drawings(els[j], removed)
                body.remove(els[j])
            j += 1
        groups = [planche_items[n:n + 9] for n in range(0, len(planche_items), 9)]
        for n, grp in reversed(list(enumerate(groups, 1))):
            say(f"Planche {n}/{len(groups)}…")
            data = make_planche(grp)
            p = e27._make("w:p")
            if ref_ppr is not None:
                p.append(deepcopy(ref_ppr))
            else:
                ppr = e27._make("w:pPr")
                ppr.append(e27._make("w:jc", val="center"))
                p.append(ppr)
            cx = PLANCHE_W_CM * EMU_PER_CM
            cy = cx * PLANCHE_PX[1] / PLANCHE_PX[0]
            p.append(e27._picture_run(d, data, cx, cy))
            anchor.addnext(p)
            planches += 1
        # « RENSEIGNEMENTS CADASTRAUX » toujours en début de page
        for el in body.iterchildren():
            if el.tag == W + "p" and ee._norm(e27._text(el)).startswith("RENSEIGNEMENTS CADASTRAUX"):
                ee._set_ppr_flag(el, "pageBreakBefore", True)
                break
    e27._drop_unused_rels(d, removed)
    target = next_version(word_path)
    say("Enregistrement…")
    temp = securite.temp_path_for(target)
    try:
        d.save(temp)
        docx.Document(temp)
    except Exception as exc:
        if os.path.exists(temp):
            os.remove(temp)  # fichier temporaire du programme uniquement
        raise ExpertiseError(f"Enregistrement impossible : {exc}") from exc
    securite.commit_temp(temp, target)
    remarques.append(f"Nouvelle version écrite à côté : « {os.path.basename(target)} » (l'original n'est pas modifié).")
    return {"path": target, "planches": planches, "remarques": remarques}


def photo_items(folder):
    """Photos du dossier « Photos » : la 1re capture .PNG (calendrier) est
    exclue ; la première photo est proposée comme façade."""
    items = e27.folder_photo_items(folder)
    for it in items:
        it["role"] = "exclure" if it["role"] == "exclure" else "planche"
    first = next((it for it in items if it["role"] == "planche"), None)
    if first:
        first["role"] = "facade"
    return items


ROLES = ("facade", "planche", "exclure")
ROLE_LABELS = {"facade": "Façade (garde)", "planche": "Planche", "exclure": "Exclure"}

