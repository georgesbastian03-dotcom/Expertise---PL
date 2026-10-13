"""Gestion des pages de plans à joindre en fin de rapport (images et/ou PDF),
et calcul automatique du nombre de pages pour la balise [[Nbre_annexes]].

Les pages sont ajoutées directement dans le .docx (une image pleine page par
page de plan, avec saut de page avant chacune) : l'export PDF (qui repart du
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


def render_pages(paths):
    """Retourne une liste ordonnée de (png_bytes, largeur_px, hauteur_px), une
    entrée par page, dans l'ordre des fichiers fournis (et des pages, pour un
    PDF). À appeler UNE SEULE FOIS par lot de génération (le même jeu de
    plans est ensuite ajouté à chaque rapport)."""
    pages = []
    for path in paths:
        ext = _ext(path)
        if ext in PDF_EXTENSIONS:
            import pymupdf as fitz  # évite l'avertissement de dépréciation de "import fitz"

            with fitz.open(path) as doc:
                for page in doc:
                    pix = page.get_pixmap(dpi=150)
                    png_bytes = pix.tobytes("png")
                    pages.append((png_bytes, pix.width, pix.height))
        else:
            with Image.open(path) as im:
                im = ImageOps.exif_transpose(im)
                png_bytes, w, h = _image_to_png_bytes(im)
                pages.append((png_bytes, w, h))
    return pages


def append_pages_to_docx(docx_path, pages):
    """Ajoute chaque page (png_bytes, w_px, h_px) comme une page pleine à la
    fin du document Word donné (une image par page, avec saut de page avant
    chacune)."""
    if not pages:
        return
    import docx
    from docx.shared import Emu, Inches

    MAX_W = Inches(6.5)
    MAX_H = Inches(9.5)

    document = docx.Document(docx_path)
    for png_bytes, w_px, h_px in pages:
        document.add_page_break()
        width = MAX_W
        height = Emu(int(width * h_px / w_px)) if w_px else MAX_H
        if height > MAX_H:
            height = MAX_H
            width = Emu(int(height * w_px / h_px)) if h_px else MAX_W
        document.add_picture(io.BytesIO(png_bytes), width=width, height=height)
    document.save(docx_path)
