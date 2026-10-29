"""État des lieux d'ENTRÉE AISBW — 100 % local.

Pour les entrées AISBW, Georges remplit le formulaire PDF (16 pages)
directement sur place. Une fois le dossier dans les dossiers définitifs :

  * le PDF du formulaire est RENOMMÉ comme les autres rapports
    (« 2026.xxxx ELLE Commune Rue n°.bte GB.pdf ») — son contenu n'est
    jamais modifié ;
  * le reportage photos est créé à partir du modèle vierge « Reportage
    photos AISBW » : en-tête avec le n° de dossier, les initiales et
    l'adresse ; photos du sous-dossier « Photos » (capture Odoo .PNG
    exclue) dans une grille alignée à 2 colonnes, 8 photos par page ;
  * le reportage est enregistré en .docx et en .pdf dans le dossier
    (dossier de référence : 2026.3529).
"""
import os
import re
import shutil
import subprocess

import entree2027 as e27
import entree_engine as ee
import securite

APP_DIR = os.path.dirname(os.path.abspath(__file__))
MODELE_REPORTAGE = os.path.join(APP_DIR, "gabarit_entree", "Reportage photos AISBW (vierge).docx")

W = e27.W
# A4, marges 2,5 cm : 16 × 24,7 cm utiles, moins le titre → 4 rangées de 2
GRILLE = dict(col_cm=8.0, box_w_cm=7.6, box_h_cm=5.7, gap_cm=0.2)

_EXCLUS_PDF = re.compile(r"\b(BAIL|OM|ORDRE|CONVENTION|MAIL|COMPTEUR|REPORTAGE|PHOTOS?)\b")


class AisbwError(Exception):
    pass


def looks_like_aisbw(folder, inv=None):
    """Indice : « AISBW » dans le nom du dossier, ou un PDF de formulaire
    sans Word dactylographié."""
    name = ee._norm(os.path.basename(os.path.normpath(folder or "")))
    if "AISBW" in name:
        return True
    inv = inv or ee.scan_folder(folder)
    rapports = [p for p in inv["docx"] if "REPORTAGE" not in ee._norm(os.path.basename(p))]
    return not rapports and bool(form_candidates(folder))


def form_candidates(folder):
    """PDF du dossier qui peuvent être le formulaire rempli sur place (les
    plus longs d'abord ; bail, OM, reportage exclus)."""
    out = []
    for f in sorted(os.listdir(folder)):
        p = os.path.join(folder, f)
        if not os.path.isfile(p) or not f.lower().endswith(".pdf") or f.startswith("~$"):
            continue
        if _EXCLUS_PDF.search(ee._norm(os.path.splitext(f)[0])):
            continue
        try:
            n = ee.page_count(p)
        except Exception:  # noqa: BLE001
            n = 0
        out.append((p, n))
    out.sort(key=lambda x: -x[1])
    return out


def final_pdf_name(folder, initiales="GB"):
    n = ee.final_docx_name(folder, initiales)
    return n[:-5] + ".pdf" if n else None


def reportage_names(folder, initiales="GB"):
    info = ee.parse_folder_name(folder)
    base = f"{info['dossier']} Reportage photos {info['reste']} {info['initiales'] or initiales}"
    base = re.sub(r'[\\/:*?"<>|]+', "-", base).strip()
    return base + ".docx", base + ".pdf"


def default_header_address(folder):
    """Adresse proposée pour l'en-tête, depuis le nom du dossier
    (« LLN Rue de la Micheline 6.202 » → « LLN, Rue de la Micheline 6/202 »)
    — à compléter (code postal, commune en toutes lettres) dans l'écran."""
    reste = ee.parse_folder_name(folder)["reste"]
    reste = re.sub(r"(\d+)\.(\d+)", r"\1/\2", reste)
    m = re.match(r"^(\S+)\s+(.*)$", reste)
    return f"{m.group(1)}, {m.group(2)}" if m else reste


def rename_form(pdf_path, new_name):
    """Renomme le formulaire (contenu inchangé). Retourne le nouveau chemin.
    Refuse d'écraser un fichier existant."""
    folder = os.path.dirname(pdf_path)
    securite.check_dossier_folder(folder)
    target = os.path.join(folder, new_name)
    if os.path.normcase(os.path.abspath(target)) == os.path.normcase(os.path.abspath(pdf_path)):
        return pdf_path
    if os.path.exists(target):
        raise AisbwError(f"« {new_name} » existe déjà : le formulaire n'a pas été renommé.")
    securite.ensure_not_locked(pdf_path)
    os.replace(pdf_path, target)
    return target


def _fill_header(document, dossier, initiales, adresse):
    for section in document.sections:
        hdr = section.header._element
        ps = list(hdr.iter(W + "p"))
        for t in hdr.iter(W + "t"):
            if re.search(r"\d{4}\.\s*(\d{3,4})?\s*/", t.text or ""):
                t.text = f"{dossier} / {initiales}"
        if adresse and len(ps) > 1:
            p = ps[1]
            if adresse not in e27._text(p):
                r = e27._make("w:r")
                t = e27._make("w:t")
                t.text = adresse
                t.set(e27.XML_SPACE, "preserve")
                r.append(t)
                p.append(r)


def export_pdf(docx_path):
    """PDF via Word (docx2pdf) si possible, sinon LibreOffice. Retourne le
    chemin du PDF ou lève une exception."""
    pdf = os.path.splitext(docx_path)[0] + ".pdf"
    try:
        from docx2pdf import convert

        convert(docx_path, pdf)
        if os.path.exists(pdf):
            return pdf
    except Exception:  # noqa: BLE001
        pass
    exe = shutil.which("soffice") or next((p for p in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe") if os.path.exists(p)), None)
    if not exe:
        raise AisbwError("PDF non créé : ni Word (docx2pdf) ni LibreOffice disponible.")
    creation = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(docx_path), docx_path],
                   capture_output=True, timeout=300, creationflags=creation)
    if not os.path.exists(pdf):
        raise AisbwError("PDF non créé (conversion LibreOffice échouée).")
    return pdf


def build_reportage(folder, items, adresse, initiales="GB", modele=None, with_pdf=True, progress_cb=None):
    """Crée le reportage photos (.docx + .pdf) dans le dossier. `items` :
    [{"path", "role", "rot"}] ; les photos « exclure » sont ignorées. Un
    reportage existant est d'abord sauvegardé hors Dropbox, puis remplacé.
    Retourne {"docx", "pdf", "photos", "remarques", "sauvegardes"}."""
    import docx

    def say(msg):
        if progress_cb:
            progress_cb(msg)

    securite.check_dossier_folder(folder)
    modele = modele or MODELE_REPORTAGE
    if not os.path.exists(modele):
        raise AisbwError(f"Modèle du reportage introuvable : {modele}")
    photos = [it for it in items if it.get("role") != "exclure"]
    if not photos:
        raise AisbwError("Aucune photo retenue pour le reportage.")
    info = ee.parse_folder_name(folder)
    name_docx, name_pdf = reportage_names(folder, initiales)
    out_docx = os.path.join(folder, name_docx)
    out_pdf = os.path.join(folder, name_pdf)
    sauvegardes, remarques = [], []
    for p in (out_docx, out_pdf):
        if os.path.exists(p):
            securite.ensure_not_locked(p)
            sauvegardes.append(securite.backup_file(p))

    d = docx.Document(modele)
    _fill_header(d, info["dossier"], initiales, adresse)
    body = d.element.body
    paras = [p for p in body.iterchildren() if p.tag == W + "p"]
    titre = next((p for p in paras if "REPORTAGE PHOTOS" in ee._norm(e27._text(p))), paras[0] if paras else None)
    grid = e27._photo_grid(d, photos, say, **GRILLE)
    if titre is not None:
        titre.addnext(grid)
    say("Enregistrement du reportage…")
    temp = securite.temp_path_for(out_docx)
    try:
        d.save(temp)
        docx.Document(temp)
    except Exception as exc:
        if os.path.exists(temp):
            os.remove(temp)  # fichier temporaire du programme uniquement
        raise AisbwError(f"Enregistrement impossible : {exc}") from exc
    securite.commit_temp(temp, out_docx)
    pdf = None
    if with_pdf:
        say("Création du PDF…")
        try:
            pdf = export_pdf(out_docx)
        except Exception as exc:  # noqa: BLE001
            remarques.append(str(exc))
    return {"docx": out_docx, "pdf": pdf, "photos": len(photos), "remarques": remarques, "sauvegardes": sauvegardes}
