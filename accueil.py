"""Écran d'accueil : grandes bulles illustrées pour naviguer entre les sections.

Chaque bulle affiche la photo de assets/accueil/ si elle existe (crédits dans
assets/accueil/CREDITS.md), sinon un dessin fait directement (Canvas tkinter).
Tout fonctionne hors ligne.
"""
import os
import sys
import tkinter as tk
from tkinter import ttk

try:
    from PIL import Image, ImageDraw, ImageTk
except ImportError:  # Pillow absent : on garde les dessins
    Image = None

# Photos des bulles : déposer par ex. « expertise.jpg » dans ce dossier pour
# remplacer le dessin de la bulle correspondante (voir les noms dans gui.py).
DOSSIER_PHOTOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "accueil")
EXTENSIONS = (".jpg", ".jpeg", ".png", ".webp")

FOND = "#f4f6f9"
TEXTE = "#1f2d3d"
TEXTE_DOUX = "#5c6570"
TRAIT = "#1f2d3d"


# --- Illustrations ---------------------------------------------------------
# Chaque fonction dessine dans un carré centré sur (cx, cy) ; u = 1 % du rayon
# de la bulle, pour que le dessin suive la taille de la bulle.

def _loupe(c, x, y, u, couleur=TRAIT):
    r = 14 * u
    c.create_oval(x - r, y - r, x + r, y + r, outline=couleur, width=max(2, 4 * u), fill="#e8f1fb")
    c.create_line(x + r * 0.7, y + r * 0.7, x + r * 1.8, y + r * 1.8,
                  fill=couleur, width=max(3, 7 * u), capstyle="round")


def _maison(c, x, y, u, mur="#ffffff", toit="#c0392b"):
    """Maison dont le bas du mur est en y."""
    w, h = 30 * u, 26 * u
    c.create_rectangle(x - w, y - h, x + w, y, fill=mur, outline=TRAIT, width=max(2, 3 * u))
    c.create_polygon(x - w - 8 * u, y - h, x, y - h - 26 * u, x + w + 8 * u, y - h,
                     fill=toit, outline=TRAIT, width=max(2, 3 * u))
    c.create_rectangle(x - 7 * u, y - 18 * u, x + 7 * u, y, fill="#8d6e63", outline=TRAIT, width=max(1, 2 * u))
    c.create_rectangle(x + 13 * u, y - 20 * u, x + 25 * u, y - 9 * u, fill="#bfe3ff", outline=TRAIT, width=max(1, 2 * u))
    c.create_rectangle(x - 25 * u, y - 20 * u, x - 13 * u, y - 9 * u, fill="#bfe3ff", outline=TRAIT, width=max(1, 2 * u))


def _fleche(c, x0, x1, y, u, couleur):
    c.create_line(x0, y, x1, y, fill=couleur, width=max(4, 8 * u), arrow="last",
                  arrowshape=(14 * u, 16 * u, 8 * u), capstyle="round")


def _document(c, x, y, u):
    w, h = 26 * u, 34 * u
    c.create_polygon(x - w, y - h, x + w - 10 * u, y - h, x + w, y - h + 10 * u, x + w, y + h, x - w, y + h,
                     fill="#ffffff", outline=TRAIT, width=max(2, 3 * u))
    for i in range(4):
        yy = y - h + 18 * u + i * 11 * u
        c.create_line(x - w + 8 * u, yy, x + w - 8 * u - (12 * u if i == 3 else 0), yy,
                      fill="#9aa5b1", width=max(1, 3 * u))


def dessin_expertise_immobiliere(c, cx, cy, u):
    _maison(c, cx - 10 * u, cy + 28 * u, u * 1.1)
    _loupe(c, cx + 26 * u, cy + 4 * u, u * 1.3)


def dessin_topographie(c, cx, cy, u):
    # Trépied
    tete = cy - 2 * u
    for dx in (-30, 0, 30):
        c.create_line(cx, tete, cx + dx * u, cy + 50 * u, fill="#8d6e63", width=max(2, 5 * u), capstyle="round")
    # Station totale
    c.create_rectangle(cx - 16 * u, tete - 30 * u, cx + 16 * u, tete, fill="#f1c40f", outline=TRAIT, width=max(2, 3 * u))
    c.create_rectangle(cx - 8 * u, tete - 24 * u, cx + 8 * u, tete - 10 * u, fill="#2c3e50", outline="")
    c.create_rectangle(cx - 26 * u, tete - 44 * u, cx + 22 * u, tete - 32 * u, fill="#34495e", outline=TRAIT, width=max(1, 2 * u))
    c.create_oval(cx + 18 * u, tete - 44 * u, cx + 30 * u, tete - 32 * u, fill="#bfe3ff", outline=TRAIT, width=max(1, 2 * u))
    # Jalon au loin
    c.create_line(cx + 46 * u, cy - 40 * u, cx + 46 * u, cy + 40 * u, fill="#c0392b", width=max(2, 4 * u), dash=(6, 4))


def dessin_entree(c, cx, cy, u):
    _maison(c, cx + 14 * u, cy + 30 * u, u)
    _fleche(c, cx - 50 * u, cx - 4 * u, cy + 18 * u, u, "#27ae60")


def dessin_sortie(c, cx, cy, u):
    _maison(c, cx - 14 * u, cy + 30 * u, u)
    _fleche(c, cx + 4 * u, cx + 50 * u, cy + 18 * u, u, "#e67e22")


def dessin_expertises(c, cx, cy, u):
    _document(c, cx - 8 * u, cy, u)
    c.create_text(cx - 8 * u, cy + 18 * u, text="€", fill="#1f6fb2", font=("Segoe UI", max(8, int(18 * u)), "bold"))
    _loupe(c, cx + 22 * u, cy + 8 * u, u)


def dessin_permis(c, cx, cy, u):
    _document(c, cx - 6 * u, cy, u)
    # Cachet « conforme »
    r = 16 * u
    x, y = cx + 22 * u, cy + 22 * u
    c.create_oval(x - r, y - r, x + r, y + r, fill="#27ae60", outline=TRAIT, width=max(1, 2 * u))
    c.create_line(x - 8 * u, y, x - 2 * u, y + 7 * u, x + 9 * u, y - 7 * u,
                  fill="#ffffff", width=max(2, 4 * u), capstyle="round", joinstyle="round")


# --- Photos ----------------------------------------------------------------

def trouver_photo(cle):
    for ext in EXTENSIONS:
        chemin = os.path.join(DOSSIER_PHOTOS, cle + ext)
        if os.path.isfile(chemin):
            return chemin
    return None


def photo_ronde(chemin, diametre):
    """Recadre la photo au centre, en carré, puis la découpe en disque lissé
    sur la couleur de fond. Retourne None si l'image est illisible."""
    if Image is None:
        return None
    try:
        img = Image.open(chemin).convert("RGB")
    except Exception:  # noqa: BLE001
        return None
    cote = min(img.size)
    gauche, haut = (img.width - cote) // 2, (img.height - cote) // 2
    img = img.crop((gauche, haut, gauche + cote, haut + cote)).resize((diametre, diametre), Image.LANCZOS)
    # Masque dessiné 4x plus grand puis réduit : bord du cercle sans escaliers.
    masque = Image.new("L", (diametre * 4, diametre * 4), 0)
    ImageDraw.Draw(masque).ellipse((0, 0, diametre * 4 - 1, diametre * 4 - 1), fill=255)
    masque = masque.resize((diametre, diametre), Image.LANCZOS)
    fond = Image.new("RGB", (diametre, diametre), FOND)
    fond.paste(img, (0, 0), masque)
    return ImageTk.PhotoImage(fond)


# --- Netteté sur les écrans agrandis ---------------------------------------

def activer_haute_resolution():
    """À appeler avant de créer la fenêtre. Sans cela, Windows affiche le
    programme en basse résolution puis l'agrandit (125 %, 150 %…) : textes et
    photos deviennent flous."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001  (Windows 7)
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:  # noqa: BLE001
        pass


def echelle(widget):
    """Facteur d'agrandissement de l'écran : 1.0 à 100 %, 1.5 à 150 %."""
    try:
        return max(1.0, widget.winfo_fpixels("1i") / 96)
    except Exception:  # noqa: BLE001
        return 1.0


# --- Bulle cliquable -------------------------------------------------------

class Bulle(tk.Canvas):
    """Grand cercle illustré, avec titre et sous-titre, qui réagit au survol."""

    def __init__(self, parent, titre, dessin, commande, couleur, sous_titre="", diametre=220, photo=None):
        # Écran à 125 % ou 150 % : on agrandit la bulle en vrais pixels plutôt
        # que de laisser Windows étirer (et flouter) une petite image.
        e = echelle(parent)
        d = round(diametre * e)
        m, h = round(10 * e), round(4 * e)
        super().__init__(parent, width=d + 2 * m, height=d + round(70 * e), bg=FOND,
                         highlightthickness=0, cursor="hand2")
        self._commande = commande
        self._couleur = couleur
        self._bord = (max(1, round(3 * e)), max(1, round(5 * e)))
        self._cercle = self.create_oval(m, h, m + d, h + d, fill=couleur, outline="#d5dbe3", width=self._bord[0])
        cx, cy = m + d / 2, h + d / 2
        chemin = trouver_photo(photo) if photo else None
        self._image = photo_ronde(chemin, d) if chemin else None
        if self._image is not None:
            self.create_image(cx, cy, image=self._image)
            self.tag_raise(self._cercle)
            self.itemconfigure(self._cercle, fill="")
        else:
            dessin(self, cx, cy, d / 200)
        self.create_text(cx, d + round(26 * e), text=titre, fill=TEXTE, font=("Segoe UI", 13, "bold"),
                         width=d + m, justify="center")
        if sous_titre:
            self.create_text(cx, d + round(50 * e), text=sous_titre, fill=TEXTE_DOUX, font=("Segoe UI", 9, "italic"))
        self.bind("<Enter>", lambda _e: self.itemconfigure(self._cercle, outline="#1f6fb2", width=self._bord[1]))
        self.bind("<Leave>", lambda _e: self.itemconfigure(self._cercle, outline="#d5dbe3", width=self._bord[0]))
        self.bind("<Button-1>", lambda _e: self._commande())


class PageBulles(tk.Frame):
    """Page avec un titre et une rangée de bulles centrées.

    `bulles` : liste de dict(titre, dessin, commande, couleur, sous_titre).
    """

    def __init__(self, parent, titre, sous_titre, bulles, diametre=220):
        super().__init__(parent, bg=FOND)
        centre = tk.Frame(self, bg=FOND)
        centre.place(relx=0.5, rely=0.45, anchor="center")
        tk.Label(centre, text=titre, bg=FOND, fg=TEXTE, font=("Segoe UI", 22, "bold")).pack(pady=(0, 4))
        if sous_titre:
            tk.Label(centre, text=sous_titre, bg=FOND, fg=TEXTE_DOUX, font=("Segoe UI", 11)).pack(pady=(0, 28))
        rangee = tk.Frame(centre, bg=FOND)
        rangee.pack()
        for b in bulles:
            Bulle(rangee, diametre=diametre, **b).pack(side="left", padx=18)


class PageModule(ttk.Frame):
    """Encadre une page existante (onglet) d'un bandeau « ← Retour » + titre.

    La page réelle n'est construite qu'à la première ouverture (`fabrique`),
    puis conservée : on retrouve sa saisie en revenant dessus.
    """

    def __init__(self, parent, titre, retour, fabrique):
        super().__init__(parent)
        bandeau = ttk.Frame(self, padding=(10, 6))
        bandeau.pack(fill="x")
        ttk.Button(bandeau, text="← Retour", command=retour).pack(side="left")
        ttk.Label(bandeau, text=titre, font=("Segoe UI", 12, "bold")).pack(side="left", padx=12)
        ttk.Separator(self).pack(fill="x")
        self._fabrique = fabrique
        self._contenu = None

    def construire(self):
        if self._contenu is None:
            self._contenu = self._fabrique(self)
            self._contenu.pack(fill="both", expand=True)


class Navigateur(tk.Frame):
    """Empile les pages et n'en affiche qu'une à la fois."""

    def __init__(self, parent):
        super().__init__(parent, bg=FOND)
        self._pages = {}

    def ajouter(self, nom, page):
        self._pages[nom] = page
        page.place(in_=self, x=0, y=0, relwidth=1, relheight=1)

    def afficher(self, nom):
        page = self._pages[nom]
        if isinstance(page, PageModule):
            page.construire()
        page.tkraise()
