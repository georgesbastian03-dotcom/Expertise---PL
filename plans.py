"""Gestion des pages de plans à joindre en fin de rapport (images et/ou PDF),
et calcul automatique du nombre de pages pour la balise [[Nbre_annexes]].

Les pages sont ajoutées directement dans le .docx (une image par plan, deux
plans courts partageant une même page quand ils y tiennent) : l'export PDF (qui repart du
.docx) les inclut donc automatiquement, sans traitement séparé.
"""
import io
import os

from PIL import Image, ImageOps

PDF_EXTENSIONS = {".pdf"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}


def _ext(path):
    return os.path.splitext(path)[1].lower()


def is_supported(path):
    return _ext(path) in PDF_EXTENSIONS or _ext(path) in IMAGE_EXTENSIONS


def count_pages(path):
    """Nombre de pages qu'occupera ce fichier une fois ajouté (1 par image,
    N par PDF). Rapide : pour un PDF, ne lit que les métadonnées."""
    ext = _ext(path)
    if ext in PDF_EXTENSIONS:
        import pymupdf as fitz  # évite l'avertissement de dépréciation de "import fitz"

        with fitz.open(path) as doc:
            return doc.page_count
    return 1


def count_total_pages(paths):
    total = 0
    for p in paths:
        try:
            total += count_pages(p)
        except Exception:  # noqa: BLE001
            # Un fichier illisible ne doit pas empêcher l'affichage du
            # compteur pour les autres — il sera de toute façon signalé
            # comme erreur au moment de la génération réelle.
            pass
    return total


def _image_to_png_bytes(pil_image):
    pil_image = pil_image.convert("RGB")
    buf = io.BytesIO()
    pil_image.save(buf, format="PNG")
    return buf.getvalue(), pil_image.width, pil_image.height


def _trim_white(pil_image, threshold=245, pad=12):
    """Retire les marges blanches autour du plan (garde une petite marge),
    pour que deux plans courts puissent tenir sur une même page."""
    rgb = pil_image.convert("RGB")
    gray = rgb.convert("L")
    mask = gray.point(lambda v: 255 if v < threshold else 0)
    box = mask.getbbox()
    if not box:
        return rgb
    l, t, r, b = box
    l, t = max(0, l - pad), max(0, t - pad)
    r, b = min(rgb.width, r + pad), min(rgb.height, b + pad)
    return rgb.crop((l, t, r, b))


def render_pages(paths):
    """Retourne une liste ordonnée de (png_bytes, largeur_px, hauteur_px), une
    entrée par page, dans l'ordre des fichiers fournis (et des pages, pour un
    PDF), marges blanches retirées. À appeler UNE SEULE FOIS par lot de
    génération (le même jeu de plans est ensuite ajouté à chaque rapport)."""
    pages = []
    for path in paths:
        ext = _ext(path)
        if ext in PDF_EXTENSIONS:
            import pymupdf as fitz  # évite l'avertissement de dépréciation de "import fitz"

            with fitz.open(path) as doc:
                for page in doc:
                    pix = page.get_pixmap(dpi=150)
                    im = Image.open(io.BytesIO(pix.tobytes("png")))
                    pages.append(_image_to_png_bytes(_trim_white(im)))
        else:
            with Image.open(path) as im:
                im = ImageOps.exif_transpose(im)
                pages.append(_image_to_png_bytes(_trim_white(im)))
    return pages


# Zone utile par défaut d'une page du rapport (pouces), marge de sécurité
# pour les paragraphes qui portent les images, et espace entre deux plans.
PAGE_W_IN = 6.5
PAGE_H_IN = 9.5
SAFETY_IN = 0.6
GAP_IN = 0.2


def usable_area(docx_path):
    """(largeur, hauteur) utiles en pouces de la dernière section du Word
    (page moins marges), ou les valeurs par défaut si illisibles."""
    try:
        import docx

        sec = docx.Document(docx_path).sections[-1]
        w = (sec.page_width - sec.left_margin - sec.right_margin) / 914400
        h = (sec.page_height - sec.top_margin - sec.bottom_margin) / 914400
        if w > 1 and h > 1:
            return min(w, PAGE_W_IN + 1), h - SAFETY_IN
    except Exception:  # noqa: BLE001
        pass
    return PAGE_W_IN, PAGE_H_IN - SAFETY_IN


def _size_in(w_px, h_px, area):
    """Taille d'affichage (pouces) d'un plan : pleine largeur, hauteur bornée."""
    page_w, page_h = area
    if not w_px or not h_px:
        return page_w, page_h
    w, h = page_w, page_w * h_px / w_px
    if h > page_h:
        h = page_h
        w = page_h * w_px / h_px
    return w, h


def pack_pages(pages, area=None):
    """Regroupe les plans par page (GB 25/11/2026) : deux plans consécutifs
    partagent une page quand ils y tiennent l'un sous l'autre à leur taille
    normale (pleine largeur), sans être réduits. Retourne une liste de
    groupes (listes de pages)."""
    area = area or (PAGE_W_IN, PAGE_H_IN - SAFETY_IN)
    groups = []
    for page in pages:
        _, w_px, h_px = page
        h = _size_in(w_px, h_px, area)[1]
        if groups and len(groups[-1]) == 1:
            prev = groups[-1][0]
            h_prev = _size_in(prev[1], prev[2], area)[1]
            if h_prev + GAP_IN + h <= area[1]:
                groups[-1].append(page)
                continue
        groups.append([page])
    return groups


def append_pages_to_docx(docx_path, pages, area=None):
    """Ajoute les plans à la fin du document Word donné : un saut de page
    avant chaque page de plans, un ou deux plans par page (voir pack_pages)."""
    if not pages:
        return
    import docx
    from docx.shared import Inches, Pt

    area = area or usable_area(docx_path)
    document = docx.Document(docx_path)
    for group in pack_pages(pages, area):
        document.add_page_break()
        for png_bytes, w_px, h_px in group:
            w, h = _size_in(w_px, h_px, area)
            document.add_picture(io.BytesIO(png_bytes), width=Inches(w), height=Inches(h))
            fmt = document.paragraphs[-1].paragraph_format
            fmt.space_before = Pt(0)
            fmt.space_after = Pt(GAP_IN * 72)
    document.save(docx_path)
