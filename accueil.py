"""Écran d'accueil : grandes bulles illustrées pour naviguer entre les sections.

Les illustrations sont dessinées directement (Canvas tkinter) : aucun fichier
image à fournir, rien à télécharger, tout fonctionne hors ligne.
"""
import tkinter as tk
from tkinter import ttk

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


# --- Bulle cliquable -------------------------------------------------------

class Bulle(tk.Canvas):
    """Grand cercle illustré, avec titre et sous-titre, qui réagit au survol."""

    def __init__(self, parent, titre, dessin, commande, couleur, sous_titre="", diametre=220):
        hauteur = diametre + 70
        super().__init__(parent, width=diametre + 20, height=hauteur, bg=FOND,
                         highlightthickness=0, cursor="hand2")
        self._commande = commande
        self._couleur = couleur
        d = diametre
        self._cercle = self.create_oval(10, 4, 10 + d, 4 + d, fill=couleur, outline="#d5dbe3", width=3)
        cx, cy = 10 + d / 2, 4 + d / 2
        dessin(self, cx, cy, d / 200)
        self.create_text(cx, d + 26, text=titre, fill=TEXTE, font=("Segoe UI", 13, "bold"),
                         width=d + 10, justify="center")
        if sous_titre:
            self.create_text(cx, d + 50, text=sous_titre, fill=TEXTE_DOUX, font=("Segoe UI", 9, "italic"))
        self.bind("<Enter>", lambda _e: self.itemconfigure(self._cercle, outline="#1f6fb2", width=5))
        self.bind("<Leave>", lambda _e: self.itemconfigure(self._cercle, outline="#d5dbe3", width=3))
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
