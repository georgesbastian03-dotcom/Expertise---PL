"""Onglets « État des lieux — entrée » et « État des lieux — sortie ».

Tout fonctionne en local, sans IA, sans internet, sans clé ni abonnement.
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk

import entree_engine as ee
import signature
import sortie_engine as se

GREY = "#5c6570"
OK_GREEN = "#1e7b34"
WARN_ORANGE = "#b35900"


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Ouverture impossible", str(exc))


def run_async(widget, work, done, failed):
    """Exécute work() dans un thread ; done/failed sont appelés sur le
    thread de l'interface (le thread de travail ne touche jamais à Tk)."""
    box = {}

    def worker():
        try:
            box["res"] = work()
        except Exception as exc:  # noqa: BLE001
            box["err"] = exc
        box["fin"] = True

    def poll():
        if not box.get("fin"):
            try:
                widget.after(120, poll)
            except tk.TclError:
                pass  # fenêtre fermée entre-temps
            return
        if "err" in box:
            failed(box["err"])
        else:
            done(box["res"])

    threading.Thread(target=worker, daemon=True).start()
    widget.after(120, poll)


def choose_signature(parent):
    p = filedialog.askopenfilename(
        parent=parent, title="Image de votre signature (PNG ou JPG)",
        filetypes=[("Images", "*.png *.jpg *.jpeg")])
    if not p:
        return False
    try:
        signature.set_signature(p)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Signature", str(exc), parent=parent)
        return False
    messagebox.showinfo("Signature", "Signature enregistrée sur ce PC (elle n'est jamais publiée).", parent=parent)
    return True


def ensure_signature(parent):
    """Avant une génération : propose de choisir la signature si aucune
    n'est enregistrée sur ce PC. Retourne True pour continuer."""
    if signature.signature_path():
        return True
    if messagebox.askyesno(
            "Signature", "Aucune image de signature n'est enregistrée sur ce PC.\n\n"
            "La choisir maintenant ? (une seule fois — « Non » pour continuer sans signature)", parent=parent):
        choose_signature(parent)
    return True


def confirm_own_folder(parent, cfg, folder):
    """Avertit si le dossier choisi porte les initiales d'un collègue."""
    mine = (cfg.get("mes_initiales") or "GB").upper()
    ini = ee.folder_initials(folder)
    if ini == mine:
        return True
    qui = f"de {ini}" if ini else "sans initiales"
    return messagebox.askyesno(
        "Dossier d'un collègue ?", f"Ce dossier est {qui}, pas {mine} :\n\n{os.path.basename(folder)}\n\n"
        "Continuer quand même ?", parent=parent)


class ScrollArea(ttk.Frame):
    """Contenu défilant (molette active quand la souris est au-dessus)."""

    def __init__(self, parent):
        super().__init__(parent)
        canvas = tk.Canvas(self, highlightthickness=0)
        vs = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.body = ttk.Frame(canvas, padding=14)
        wid = canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))

        def wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)) or (-1 if e.delta > 0 else 1), "units")

        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
        self.canvas = canvas


def section(parent, title):
    f = ttk.LabelFrame(parent, text=title, padding=10)
    f.pack(fill="x", pady=(0, 10))
    return f


def labeled_entry(parent, row, col, label, var, width=28):
    ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=(0, 6), pady=2)
    e = ttk.Entry(parent, textvariable=var, width=width)
    e.grid(row=row, column=col + 1, sticky="we", padx=(0, 14), pady=2)
    return e


# ==========================================================================
#  SORTIE
# ==========================================================================
class PosteDialog(tk.Toplevel):
    """Ajout / modification d'une ligne « Autres postes »."""

    def __init__(self, parent, item=None, kind="poste"):
        super().__init__(parent)
        self.result = None
        self.kind = (item or {}).get("type", kind)
        self.title("Pièce (titre)" if self.kind == "titre" else "Poste de dégât")
        self.transient(parent)
        self.resizable(True, False)
        item = item or {}
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        self.vars = {k: tk.StringVar(value=str(item.get(k, "") or "")) for k in
                     ("texte", "qte", "nature", "pu", "usure", "tva")}
        if not self.vars["tva"].get():
            self.vars["tva"].set("6")
        if not self.vars["qte"].get() and self.kind == "poste":
            self.vars["qte"].set("1")
        if self.kind == "titre":
            ttk.Label(frm, text="Nom de la pièce (ex. « Séjour », « 3 chambres, hall et séjour ») :").pack(anchor="w")
            e = ttk.Entry(frm, textvariable=self.vars["texte"], width=60)
            e.pack(fill="x", pady=(4, 10))
        else:
            ttk.Label(frm, text="Description du poste (recopiée du manuscrit) :").pack(anchor="w")
            self.txt = tk.Text(frm, height=4, width=70, wrap="word", font=("Segoe UI", 10))
            self.txt.insert("1.0", item.get("texte", ""))
            self.txt.pack(fill="x", pady=(4, 8))
            e = self.txt
            g = ttk.Frame(frm)
            g.pack(fill="x")
            labeled_entry(g, 0, 0, "Quantité", self.vars["qte"], 8)
            labeled_entry(g, 0, 2, "Nature / surface", self.vars["nature"], 30)
            labeled_entry(g, 1, 0, "PU (€)", self.vars["pu"], 14)
            labeled_entry(g, 1, 2, "Usure déduite (%)", self.vars["usure"], 8)
            ttk.Label(g, text="TVA").grid(row=2, column=0, sticky="w", pady=2)
            cb = ttk.Combobox(g, textvariable=self.vars["tva"], values=["6", "21", "0"], width=6, state="readonly")
            cb.grid(row=2, column=1, sticky="w", pady=2)
            ttk.Label(g, text="PU : un montant (150) ou un calcul (80*15+3*70). Vide = prix pas encore fixé.",
                      foreground=GREY).grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 0))
        b = ttk.Frame(frm)
        b.pack(fill="x", pady=(10, 0))
        ttk.Button(b, text="Annuler", command=self.destroy).pack(side="right")
        ttk.Button(b, text="OK", command=self._ok).pack(side="right", padx=6)
        self.bind("<Escape>", lambda _e: self.destroy())
        e.focus_set()
        self.grab_set()

    def _ok(self):
        d = {k: v.get().strip() for k, v in self.vars.items()}
        d["type"] = self.kind
        if self.kind == "poste":
            d["texte"] = self.txt.get("1.0", "end").strip()
            if d["pu"]:
                try:
                    se.eval_montant(d["pu"])
                except Exception:  # noqa: BLE001
                    messagebox.showerror("Prix illisible", f"Le prix « {d['pu']} » n'est pas compris.", parent=self)
                    return
        if not d["texte"]:
            messagebox.showerror("Texte manquant", "Le texte est vide.", parent=self)
            return
        self.result = d
        self.destroy()


class SortieTab(ttk.Frame):
    def __init__(self, parent, cfg, save_cfg):
        super().__init__(parent)
        self.cfg = cfg
        self.save_cfg = save_cfg
        self.autres = []
        area = ScrollArea(self)
        area.pack(fill="both", expand=True)
        self.root_frame = area.body
        self.v = {}
        self._build()
        self._load_data(se.empty_data())

    # -- helpers -----------------------------------------------------------
    def sv(self, key, value=""):
        var = tk.StringVar(value=value)
        var.trace_add("write", lambda *_a: self._schedule_total())
        self.v[key] = var
        return var

    def _build(self):
        R = self.root_frame
        ttk.Label(
            R, wraplength=780, justify="left", foreground=GREY,
            text=("Vous lisez le document « valorisations manuscrites » et saisissez les valeurs ci-dessous. "
                  "Le programme calcule le total comme le tableau Excel, remplit une copie du tableau type "
                  "(aucune ligne supprimée, mise en page inchangée, signature au-dessus de « po L'expert en charge ») "
                  "et prépare le PV Word avec le montant en chiffres et en lettres. 100 % local, sans IA."),
        ).pack(anchor="w", pady=(0, 10))

        # Dossier
        f = section(R, "Dossier")
        f.columnconfigure(1, weight=1)
        self.folder_var = tk.StringVar()
        ttk.Label(f, text="Dossier du n° :").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.folder_var).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(f, text="Parcourir…", command=self._choose_folder).grid(row=0, column=2)
        g = ttk.Frame(f)
        g.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))
        labeled_entry(g, 0, 0, "N° de dossier", self.sv("dossier"), 12)
        labeled_entry(g, 0, 2, "Initiales", self.sv("initiales", "GB"), 6)
        ttk.Button(g, text="Ouvrir une saisie…", command=self._open_saisie).grid(row=0, column=4, padx=(10, 4))
        ttk.Button(g, text="Enregistrer la saisie", command=self._save_saisie).grid(row=0, column=5)

        # Bien
        f = section(R, "Adresse du bien")
        labeled_entry(f, 0, 0, "Rue et n°", self.sv("bien.rue"), 40)
        labeled_entry(f, 0, 2, "CP", self.sv("bien.cp"), 8)
        labeled_entry(f, 0, 4, "Commune", self.sv("bien.commune"), 24)

        # Parties
        f = section(R, "Parties (noms sur une ligne ; si trop long, abréger le prénom : « Bastian G. »)")
        for col, (key, title) in enumerate((("locataire", "Locataire(s)"), ("bailleur", "Bailleur(s)"))):
            p = ttk.Frame(f)
            p.grid(row=0, column=col, sticky="nw", padx=(0, 20))
            ttk.Label(p, text=title, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
            labeled_entry(p, 1, 0, "Nom (Excel)", self.sv(f"{key}.nom_excel"), 34)
            labeled_entry(p, 2, 0, "Nom (PV)", self.sv(f"{key}.nom_pv"), 34)
            ttk.Button(p, text="Proposer", command=lambda k=key: self._propose_nom_pv(k)).grid(row=2, column=2)
            labeled_entry(p, 3, 0, "Rue et n°", self.sv(f"{key}.rue"), 34)
            labeled_entry(p, 4, 0, "CP + commune", self.sv(f"{key}.cpville"), 34)
            labeled_entry(p, 5, 0, "Représenté par", self.sv(f"{key}.repr"), 34)
        ttk.Label(f, foreground=GREY, text="Excel : « M. DU TOIN Jonathan & Mme SINGH Kiara » — PV : « Jonathan DU TOIN & Kiara SINGH » "
                  "(« Proposer » remplit la forme PV). « Les précités » devient « Représenté par le précité » dans le PV.", wraplength=760, justify="left").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Dates
        f = section(R, "Dates")
        labeled_entry(f, 0, 0, "État des lieux d'entrée du", self.sv("date_entree"), 12)
        labeled_entry(f, 0, 2, "Sortie du", self.sv("date_sortie"), 12)
        labeled_entry(f, 0, 4, "Clôture (à Wavre, le)", self.sv("date_cloture", se.today_str()), 12)

        # Compteurs
        f = section(R, "Relevés des compteurs (vérifier avec l'ELLE et les relevés du dossier)")
        for c, h in enumerate(("", "N° compteur", "Index 1", "Index 2 / autres")):
            ttk.Label(f, text=h, foreground=GREY).grid(row=0, column=c, sticky="w")
        for r, (key, label) in enumerate((("eau", "Eau"), ("elec", "Électricité"), ("gaz", "Gaz"), ("chaleur", "Chaleur (S/N)")), 1):
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8))
            for c, sub in enumerate(("num", "index1", "index2"), 1):
                ttk.Entry(f, textvariable=self.sv(f"compteurs.{key}.{sub}"), width=26).grid(row=r, column=c, padx=(0, 8), pady=2)
        ttk.Label(f, foreground=GREY, text="Unité ajoutée automatiquement si vous ne tapez qu'un nombre (m³ eau/gaz, KWh élec/chaleur).").grid(
            row=5, column=0, columnspan=4, sticky="w")
        g = ttk.Frame(f)
        g.grid(row=6, column=0, columnspan=4, sticky="w", pady=(6, 0))
        labeled_entry(g, 0, 0, "Citerne : litres", self.sv("citerne.litres"), 8)
        labeled_entry(g, 0, 2, "prix / litre (€)", self.sv("citerne.prix"), 8)

        # Clés / entretiens
        f = section(R, "Clés, détecteurs et documents d'entretien (quantités ; tarifs du tableau type)")
        r = 0
        for k, lab in se.CLES_LABELS.items():
            labeled_entry(f, r // 2, (r % 2) * 2, lab, self.sv(f"cles.{k}"), 6)
            r += 1
        r = 4
        for k, lab in se.ENTRETIEN_LABELS.items():
            labeled_entry(f, r // 2, (r % 2) * 2, lab, self.sv(f"entretiens.{k}"), 6)
            r += 1

        # Nettoyage
        f = section(R, "Défauts de nettoyage (TVA 21 %)")
        for c, h in enumerate(("Texte", "Qté", "Unité", "PU (€)")):
            ttk.Label(f, text=h, foreground=GREY).grid(row=0, column=c, sticky="w")
        for i in range(len(se.NETTOYAGE_ROWS)):
            ttk.Entry(f, textvariable=self.sv(f"nett.{i}.texte"), width=60).grid(row=i + 1, column=0, padx=(0, 6), pady=1)
            ttk.Entry(f, textvariable=self.sv(f"nett.{i}.qte"), width=6).grid(row=i + 1, column=1, padx=(0, 6))
            ttk.Entry(f, textvariable=self.sv(f"nett.{i}.unite"), width=14).grid(row=i + 1, column=2, padx=(0, 6))
            ttk.Entry(f, textvariable=self.sv(f"nett.{i}.pu"), width=10).grid(row=i + 1, column=3)

        # Petites interventions
        f = section(R, "Petites interventions locatives (TVA 21 %)")
        labeled_entry(f, 0, 0, "Main d'œuvre (heures)", self.sv("mo.qte"), 6)
        labeled_entry(f, 0, 2, "PU (€/h)", self.sv("mo.pu", "50"), 8)
        g = ttk.Frame(f)
        g.grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        for c, h in enumerate(("Petites fournitures — texte", "Qté", "PU (€)", "Usure %")):
            ttk.Label(g, text=h, foreground=GREY).grid(row=0, column=c, sticky="w")
        for i in range(len(se.FOURNITURES_ROWS)):
            ttk.Entry(g, textvariable=self.sv(f"four.{i}.texte"), width=60).grid(row=i + 1, column=0, padx=(0, 6), pady=1)
            ttk.Entry(g, textvariable=self.sv(f"four.{i}.qte"), width=6).grid(row=i + 1, column=1, padx=(0, 6))
            ttk.Entry(g, textvariable=self.sv(f"four.{i}.pu"), width=10).grid(row=i + 1, column=2, padx=(0, 6))
            ttk.Entry(g, textvariable=self.sv(f"four.{i}.usure"), width=6).grid(row=i + 1, column=3)

        # Autres postes
        f = section(R, "Autres postes (dans l'ordre du manuscrit : une ligne « pièce » puis ses postes)")
        cols = ("texte", "qte", "nature", "pu", "usure", "tva", "total")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=9)
        for c, h, w in (("texte", "Pièce / poste", 380), ("qte", "Qté", 45), ("nature", "Nature / surface", 160),
                        ("pu", "PU", 90), ("usure", "Usure", 55), ("tva", "TVA", 45), ("total", "Total", 80)):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w" if c in ("texte", "nature") else "e")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.tag_configure("titre", font=("Segoe UI", 10, "bold"))
        self.tree.bind("<Double-1>", lambda _e: self._edit_poste())
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        f.columnconfigure(0, weight=1)
        b = ttk.Frame(f)
        b.grid(row=1, column=0, sticky="w", pady=(6, 0))
        for txt, cmd in (("+ Pièce", lambda: self._add_poste("titre")), ("+ Poste", lambda: self._add_poste("poste")),
                         ("Modifier", self._edit_poste), ("Supprimer", self._del_poste),
                         ("Monter", lambda: self._move_poste(-1)), ("Descendre", lambda: self._move_poste(1))):
            ttk.Button(b, text=txt, command=cmd).pack(side="left", padx=(0, 4))

        # Récapitulatif
        f = section(R, "Récapitulatif")
        labeled_entry(f, 0, 0, "Autre suivant accord entre parties (€)", self.sv("autre_accord"), 10)
        labeled_entry(f, 0, 2, "Indisponibilité des lieux (texte)", self.sv("indisponibilite", "non inclus - estimé à "), 30)
        self.coord_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Coordination travaux (10 %)", variable=self.coord_var,
                        command=self._schedule_total).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.total_var = tk.StringVar(value="Total : 0,00 €")
        ttk.Label(f, textvariable=self.total_var, font=("Segoe UI", 11, "bold")).grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))
        self.total_detail = tk.StringVar()
        ttk.Label(f, textvariable=self.total_detail, foreground=GREY).grid(row=3, column=0, columnspan=4, sticky="w")

        f = ttk.Frame(R)
        f.pack(fill="x", pady=(4, 10))
        ttk.Button(f, text="Générer l'Excel et le PV Word dans le dossier", command=self._generate).pack(side="left")
        self.open_btn = ttk.Button(f, text="Ouvrir le dossier", command=self._open_folder, state="disabled")
        self.open_btn.pack(side="left", padx=8)
        self.status_var = tk.StringVar()
        ttk.Label(R, textvariable=self.status_var, foreground=GREY, wraplength=780, justify="left").pack(anchor="w")

    # -- données -----------------------------------------------------------
    def _get_data(self):
        g = lambda k: self.v[k].get().strip()  # noqa: E731
        d = se.empty_data()
        d["dossier"], d["initiales"] = g("dossier"), g("initiales") or "GB"
        d["dossier_folder"] = self.folder_var.get().strip()
        for k in ("rue", "cp", "commune"):
            d["bien"][k] = g(f"bien.{k}")
        for p in ("locataire", "bailleur"):
            for k in ("nom_excel", "nom_pv", "rue", "cpville", "repr"):
                d[p][k] = g(f"{p}.{k}")
        d["date_entree"], d["date_sortie"], d["date_cloture"] = g("date_entree"), g("date_sortie"), g("date_cloture")
        for c in d["compteurs"]:
            for s in ("num", "index1", "index2"):
                d["compteurs"][c][s] = g(f"compteurs.{c}.{s}")
        d["citerne"] = {"litres": g("citerne.litres"), "prix": g("citerne.prix")}
        d["cles"] = {k: g(f"cles.{k}") for k in se.CLES_ROWS}
        d["entretiens"] = {k: g(f"entretiens.{k}") for k in se.ENTRETIEN_ROWS}
        d["nettoyage"] = []
        for i in range(len(se.NETTOYAGE_ROWS)):
            line = {k: g(f"nett.{i}.{k}") for k in ("texte", "qte", "unite", "pu")}
            if line["texte"] or line["qte"]:
                d["nettoyage"].append(line)
        d["main_oeuvre"] = {"qte": g("mo.qte"), "pu": g("mo.pu")}
        d["fournitures"] = []
        lines = [{k: g(f"four.{i}.{k}") for k in ("texte", "qte", "pu", "usure")} for i in range(len(se.FOURNITURES_ROWS))]
        while lines and not any(lines[-1].values()):
            lines.pop()
        d["fournitures"] = lines
        d["autres"] = [dict(a) for a in self.autres]
        d["autre_accord"], d["indisponibilite"] = g("autre_accord"), self.v["indisponibilite"].get()
        d["coordination"] = bool(self.coord_var.get())
        return d

    def _load_data(self, d):
        d = se.merge_defaults(d)
        self._loading = True
        s = lambda k, val: self.v[k].set("" if val is None else str(val))  # noqa: E731
        s("dossier", d["dossier"])
        s("initiales", d["initiales"])
        self.folder_var.set(d.get("dossier_folder", ""))
        for k in ("rue", "cp", "commune"):
            s(f"bien.{k}", d["bien"][k])
        for p in ("locataire", "bailleur"):
            for k in ("nom_excel", "nom_pv", "rue", "cpville", "repr"):
                s(f"{p}.{k}", d[p].get(k, ""))
        s("date_entree", d["date_entree"])
        s("date_sortie", d["date_sortie"])
        s("date_cloture", d.get("date_cloture") or se.today_str())
        for c in d["compteurs"]:
            for x in ("num", "index1", "index2"):
                s(f"compteurs.{c}.{x}", d["compteurs"][c].get(x, ""))
        s("citerne.litres", d["citerne"].get("litres"))
        s("citerne.prix", d["citerne"].get("prix"))
        for k in se.CLES_ROWS:
            s(f"cles.{k}", d["cles"].get(k))
        for k in se.ENTRETIEN_ROWS:
            s(f"entretiens.{k}", d["entretiens"].get(k))
        for i in range(len(se.NETTOYAGE_ROWS)):
            line = d["nettoyage"][i] if i < len(d["nettoyage"]) else {}
            for k in ("texte", "qte", "unite", "pu"):
                s(f"nett.{i}.{k}", line.get(k, ""))
        s("mo.qte", d["main_oeuvre"].get("qte"))
        s("mo.pu", d["main_oeuvre"].get("pu") or "50")
        for i in range(len(se.FOURNITURES_ROWS)):
            line = d["fournitures"][i] if i < len(d["fournitures"]) else {}
            for k in ("texte", "qte", "pu", "usure"):
                s(f"four.{i}.{k}", line.get(k, ""))
        self.autres = [dict(a) for a in d["autres"]]
        s("autre_accord", d.get("autre_accord"))
        s("indisponibilite", d.get("indisponibilite") or "non inclus - estimé à ")
        self.coord_var.set(bool(d.get("coordination")))
        self._loading = False
        self._refresh_tree()
        self._update_total()

    # -- actions -------------------------------------------------------------
    def _choose_folder(self):
        start = self.cfg.get("last_sortie_parent", "") or None
        folder = filedialog.askdirectory(title="Dossier du n° (sortie locative)", initialdir=start)
        if not folder:
            return
        try:
            ee.securite.check_dossier_folder(folder)
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning("Dossier", str(exc))
            return
        if not confirm_own_folder(self, self.cfg, folder):
            return
        self.cfg["last_sortie_parent"] = os.path.dirname(folder)
        self.save_cfg()
        self.folder_var.set(folder)
        info = se.parse_folder_name(folder)
        saved = os.path.join(folder, f"{info['dossier']} saisie sortie.json") if info["dossier"] else None
        if saved and os.path.exists(saved) and messagebox.askyesno(
                "Saisie existante", "Une saisie existe déjà pour ce dossier. La rouvrir ?"):
            try:
                d = se.load_saisie(saved)
                d["dossier_folder"] = folder
                self._load_data(d)
                self.status_var.set(f"Saisie rouverte : {os.path.basename(saved)}")
                return
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Lecture impossible", str(exc))
        if info["dossier"]:
            self.v["dossier"].set(info["dossier"])
        if info["initiales"]:
            self.v["initiales"].set(info["initiales"])

    def _open_saisie(self):
        p = filedialog.askopenfilename(title="Ouvrir une saisie", filetypes=[("Saisie sortie", "*.json")],
                                       initialdir=self.folder_var.get() or None)
        if not p:
            return
        try:
            d = se.load_saisie(p)
            if not d.get("dossier_folder") or not os.path.isdir(d["dossier_folder"]):
                d["dossier_folder"] = os.path.dirname(p)
            self._load_data(d)
            self.status_var.set(f"Saisie ouverte : {os.path.basename(p)}")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Lecture impossible", str(exc))

    def _save_saisie(self):
        d = self._get_data()
        folder = d["dossier_folder"]
        if not folder or not os.path.isdir(folder) or not d["dossier"]:
            messagebox.showwarning("Dossier", "Choisissez d'abord le dossier du n° et le n° de dossier.")
            return
        p = os.path.join(folder, f"{d['dossier']} saisie sortie.json")
        se.save_saisie(d, p)
        self.status_var.set(f"Saisie enregistrée : {p}")

    def _propose_nom_pv(self, key):
        self.v[f"{key}.nom_pv"].set(se.nom_pv_depuis_excel(self.v[f"{key}.nom_excel"].get()))

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, a in enumerate(self.autres):
            if a.get("type") == "titre":
                self.tree.insert("", "end", iid=str(i), values=(a.get("texte", ""), "", "", "", "", "", ""), tags=("titre",))
            else:
                try:
                    pu = se.eval_montant(a.get("pu")) or 0.0
                    tot = se.to_num(a.get("qte")) * pu * (1 - se.to_num(a.get("usure")) / 100)
                    tot_s = f"{tot:,.2f} €".replace(",", " ").replace(".", ",")
                except Exception:  # noqa: BLE001
                    tot_s = "?"
                self.tree.insert("", "end", iid=str(i), values=(
                    "   - " + (a.get("texte", "").replace("\n", " ")), a.get("qte", ""), a.get("nature", ""),
                    a.get("pu", "") or "(à fixer)", (a.get("usure") or "0") + " %", (a.get("tva") or "6") + " %", tot_s))

    def _selected_index(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _add_poste(self, kind):
        dlg = PosteDialog(self, kind=kind)
        self.wait_window(dlg)
        if dlg.result:
            i = self._selected_index()
            pos = len(self.autres) if i is None else i + 1
            if len(self.autres) >= len(se.AUTRES_ROWS):
                messagebox.showwarning("Tableau plein", f"Le tableau type prévoit {len(se.AUTRES_ROWS)} lignes au maximum.")
                return
            self.autres.insert(pos, dlg.result)
            self._refresh_tree()
            self.tree.selection_set(str(pos))
            self._update_total()

    def _edit_poste(self):
        i = self._selected_index()
        if i is None:
            return
        dlg = PosteDialog(self, item=self.autres[i])
        self.wait_window(dlg)
        if dlg.result:
            self.autres[i] = dlg.result
            self._refresh_tree()
            self.tree.selection_set(str(i))
            self._update_total()

    def _del_poste(self):
        i = self._selected_index()
        if i is None:
            return
        del self.autres[i]
        self._refresh_tree()
        self._update_total()

    def _move_poste(self, delta):
        i = self._selected_index()
        if i is None:
            return
        j = i + delta
        if 0 <= j < len(self.autres):
            self.autres[i], self.autres[j] = self.autres[j], self.autres[i]
            self._refresh_tree()
            self.tree.selection_set(str(j))

    def _schedule_total(self):
        if getattr(self, "_loading", False):
            return
        if getattr(self, "_total_job", None):
            self.after_cancel(self._total_job)
        self._total_job = self.after(400, self._update_total)

    def _update_total(self):
        self._total_job = None
        try:
            c = se.compute(self._get_data())
        except Exception as exc:  # noqa: BLE001
            self.total_var.set("Total : calcul impossible")
            self.total_detail.set(str(exc))
            return
        fmt = lambda x: se.montant_lettres.montant_en_chiffres(x)  # noqa: E731
        self.total_var.set(f"Total indemnité : {fmt(c['total'])}")
        self.total_detail.set(f"HTVA {fmt(c['total_htva'])} — TVA 6 % {fmt(c['tva6'])} — TVA 21 % {fmt(c['tva21'])}"
                              + (f" — coordination {fmt(c['coordination'])}" if c["coordination"] else "")
                              + (f" — autre {fmt(c['autre'])}" if c["autre"] else "")
                              + f"\n{se.montant_lettres.montant_en_lettres(c['total'])}")

    def _generate(self):
        d = self._get_data()
        missing = [lab for lab, val in (("n° de dossier", d["dossier"]), ("dossier du n°", d["dossier_folder"]),
                                        ("rue du bien", d["bien"]["rue"]), ("commune du bien", d["bien"]["commune"]),
                                        ("date de sortie", d["date_sortie"])) if not val]
        if missing:
            messagebox.showwarning("Champs manquants", "À compléter : " + ", ".join(missing) + ".")
            return
        exist = se.existing_outputs(d)
        if exist and not messagebox.askyesno(
                "Fichiers existants",
                "Ces fichiers existent déjà dans le dossier :\n\n" + "\n".join(os.path.basename(p) for p in exist)
                + "\n\nLes remplacer ? (une copie de sauvegarde est faite d'abord, hors Dropbox)"):
            return
        ensure_signature(self)
        try:
            res = se.generate(d, overwrite=bool(exist))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Génération impossible", str(exc))
            return
        self.open_btn.configure(state="normal")
        self.status_var.set(f"Créés dans le dossier :\n• {os.path.basename(res['excel'])}\n• {os.path.basename(res['pv'])}\n"
                            f"Total : {se.montant_lettres.montant_en_chiffres(res['total'])}")
        if messagebox.askyesno("Terminé", "L'Excel et le PV Word ont été créés dans le dossier du n°.\n\n"
                               "Ouvrir l'Excel pour vérifier ?"):
            open_path(res["excel"])

    def _open_folder(self):
        if self.folder_var.get():
            open_path(self.folder_var.get())


# ==========================================================================
#  ENTRÉE
# ==========================================================================
class PhotoPicker(tk.Toplevel):
    """Grille de vignettes avec cases à cocher (ordre conservé)."""

    COLS = 5

    def __init__(self, parent, items, on_done):
        super().__init__(parent)
        self.title("Photos à injecter dans le rapport")
        self.geometry("980x720")
        self.transient(parent)
        self.on_done = on_done
        self.items = items  # [(path, checked, note)]
        self.vars = []
        self._imgs = []
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, wraplength=900, justify="left", foreground=GREY, text=(
            "Décochez la capture d'écran du n° de dossier et les photos d'index compteurs. Gardez la photo des clés. "
            "L'ordre est celui des numéros de photo.")).pack(anchor="w")
        self.count_var = tk.StringVar()
        ttk.Label(top, textvariable=self.count_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(4, 0))
        area = ScrollArea(self)
        area.pack(fill="both", expand=True)
        self.grid_frame = area.body
        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="x")
        ttk.Button(bot, text="Valider la sélection", command=self._done).pack(side="right")
        ttk.Button(bot, text="Tout cocher", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(bot, text="Tout décocher", command=lambda: self._set_all(False)).pack(side="left", padx=6)
        for i, (path, checked, note) in enumerate(items):
            var = tk.BooleanVar(value=checked)
            var.trace_add("write", lambda *_a: self._count())
            self.vars.append(var)
            cell = ttk.Frame(self.grid_frame, padding=4)
            cell.grid(row=i // self.COLS, column=i % self.COLS, sticky="n")
            lbl = ttk.Label(cell, text="…")
            lbl.pack()
            ttk.Checkbutton(cell, text=os.path.basename(path), variable=var).pack()
            if note:
                ttk.Label(cell, text=note, foreground=WARN_ORANGE, wraplength=160, justify="center").pack()
            lbl.bind("<Button-1>", lambda _e, v=var: v.set(not v.get()))
            lbl.bind("<Double-1>", lambda _e, p=path: open_path(p))
            self._load_thumb_async(path, lbl)
        self._count()
        self.grab_set()

    def _load_thumb_async(self, path, lbl):
        from PIL import ImageTk

        def work():
            return ee.thumbnail(path)

        def done(pil):
            try:
                img = ImageTk.PhotoImage(pil)
                self._imgs.append(img)
                lbl.configure(image=img, text="")
            except Exception:  # noqa: BLE001
                pass

        run_async(self, work, done, lambda _e: lbl.configure(text="(aperçu impossible)"))

    def _count(self):
        n = sum(v.get() for v in self.vars)
        self.count_var.set(f"{n} photo(s) retenue(s) sur {len(self.vars)}")

    def _set_all(self, val):
        for v in self.vars:
            v.set(val)

    def _done(self):
        selected = [(p, v.get(), note) for (p, _c, note), v in zip(self.items, self.vars)]
        self.on_done(selected)
        self.destroy()


class DocViewer(tk.Toplevel):
    """Visionneuse du bail / de l'OM, pour contrôler les coordonnées."""

    def __init__(self, parent, path):
        super().__init__(parent)
        self.title(os.path.basename(path))
        self.geometry("900x1000")
        self.path = path
        self.page = 0
        self.zoom = 1.0
        try:
            self.n = ee.page_count(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Lecture impossible", str(exc), parent=parent)
            self.destroy()
            return
        bar = ttk.Frame(self, padding=6)
        bar.pack(fill="x")
        ttk.Button(bar, text="◀", width=3, command=lambda: self._go(-1)).pack(side="left")
        self.lab = ttk.Label(bar, text="")
        self.lab.pack(side="left", padx=8)
        ttk.Button(bar, text="▶", width=3, command=lambda: self._go(1)).pack(side="left")
        ttk.Button(bar, text="Zoom +", command=lambda: self._zoom(1.25)).pack(side="left", padx=(16, 4))
        ttk.Button(bar, text="Zoom −", command=lambda: self._zoom(0.8)).pack(side="left")
        ttk.Button(bar, text="Ouvrir le fichier", command=lambda: open_path(path)).pack(side="right")
        self.canvas = tk.Canvas(self, background="#888")
        xs = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        ys = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=xs.set, yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom", fill="x")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120) or 1, "units"))
        self._show()

    def _go(self, d):
        if 0 <= self.page + d < self.n:
            self.page += d
            self._show()

    def _zoom(self, f):
        self.zoom = max(0.4, min(4.0, self.zoom * f))
        self._show()

    def _show(self):
        from PIL import ImageTk

        self.lab.configure(text=f"Page {self.page + 1} / {self.n}")
        try:
            img = ee.render_pdf_page(self.path, self.page, dpi=int(100 * self.zoom))
        except Exception as exc:  # noqa: BLE001
            self.canvas.delete("all")
            self.canvas.create_text(20, 20, anchor="nw", text=f"Aperçu impossible : {exc}", fill="white")
            return
        if os.path.splitext(self.path)[1].lower() != ".pdf":
            w = int(850 * self.zoom)
            img = img.resize((w, int(img.height * w / img.width)))
        self._img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self._img)
        self.canvas.configure(scrollregion=(0, 0, img.width, img.height))
        self.canvas.yview_moveto(0)


class EntreeTab(ttk.Frame):
    def __init__(self, parent, cfg, save_cfg):
        super().__init__(parent)
        self.cfg = cfg
        self.save_cfg = save_cfg
        self.folder = None
        self.photo_items = []      # [(path, checked, note)]
        self.parties = []          # read_parties()
        self.party_widgets = []    # [(Text, BooleanVar, status_var)]
        self.doc_results = {}
        area = ScrollArea(self)
        area.pack(fill="both", expand=True)
        self.R = area.body
        self._pmsg = None
        self._build()
        self._tick()

    def _tick(self):
        if self._pmsg is not None:
            self.progress_var.set(self._pmsg)
            self._pmsg = None
        try:
            self.after(200, self._tick)
        except tk.TclError:
            pass

    def _build(self):
        R = self.R
        ttk.Label(R, wraplength=780, justify="left", foreground=GREY, text=(
            "Finalise le Word renvoyé par le bureau de dactylographie : photos rangées dans « Photos » et injectées "
            "2 par page sous le titre PHOTOGRAPHIES, coordonnées des parties contrôlées avec le bail et l'OM, "
            "date de clôture du jour et signature, puis renommage. Le texte dicté n'est jamais modifié "
            "(les repères « 12*42 » restent). Le Word d'origine est écrasé, pas de PDF.")).pack(anchor="w", pady=(0, 10))

        f = section(R, "1. Dossier")
        f.columnconfigure(1, weight=1)
        self.folder_var = tk.StringVar()
        ttk.Label(f, text="Dossier du n° :").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.folder_var).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(f, text="Parcourir…", command=self._choose_folder).grid(row=0, column=2)
        ttk.Button(f, text="Plusieurs dossiers…", command=lambda: EntreeBatchWindow(self, self.cfg, self.save_cfg)).grid(
            row=0, column=3, padx=(6, 0))
        ttk.Label(f, text="Word à finaliser :").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.docx_var = tk.StringVar()
        self.docx_combo = ttk.Combobox(f, textvariable=self.docx_var, state="readonly")
        self.docx_combo.grid(row=1, column=1, sticky="we", padx=6, pady=(6, 0))
        self.docx_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_parties())
        self.newname_var = tk.StringVar()
        ttk.Label(f, text="Nom final :").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(f, textvariable=self.newname_var).grid(row=2, column=1, sticky="we", padx=6, pady=(6, 0))
        self.info_var = tk.StringVar()
        ttk.Label(f, textvariable=self.info_var, foreground=GREY, wraplength=760, justify="left").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))

        f = section(R, "2. Photos")
        self.photo_var = tk.StringVar(value="—")
        ttk.Label(f, textvariable=self.photo_var, wraplength=600, justify="left").pack(side="left")
        ttk.Button(f, text="Choisir les photos…", command=self._pick_photos).pack(side="right")
        ttk.Button(f, text="Ranger les photos en vrac", command=self._ranger).pack(side="right", padx=6)

        f = section(R, "3. Coordonnées des parties (bail d'abord, puis OM ; différence → surligner en jaune)")
        self.parties_frame = ttk.Frame(f)
        self.parties_frame.pack(fill="x")
        ttk.Label(self.parties_frame, text="Choisissez d'abord le dossier.", foreground=GREY).pack(anchor="w")
        g = ttk.Frame(f)
        g.pack(fill="x", pady=(8, 0))
        ttk.Label(g, text="Documents :").pack(side="left")
        self.doc_var = tk.StringVar()
        self.doc_combo = ttk.Combobox(g, textvariable=self.doc_var, state="readonly", width=50)
        self.doc_combo.pack(side="left", padx=6)
        ttk.Button(g, text="Voir", command=self._view_doc).pack(side="left")
        ttk.Button(g, text="Ajouter un document…", command=self._add_doc).pack(side="left", padx=6)
        ttk.Button(g, text="Lire le bail / l'OM", command=self._read_docs).pack(side="left")
        self.ocr_var = tk.StringVar()
        ttk.Label(f, textvariable=self.ocr_var, foreground=GREY, wraplength=760, justify="left").pack(anchor="w", pady=(6, 0))
        self.found = tk.Text(f, height=8, wrap="word", font=("Segoe UI", 9))
        self.found.pack(fill="x", pady=(6, 0))
        self.found.insert("1.0", "Les noms et lignes trouvés dans le bail / l'OM s'afficheront ici.")
        self.found.configure(state="disabled")
        self._refresh_ocr_status()

        f = section(R, "4. Clôture")
        self.date_var = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        labeled_entry(f, 0, 0, "Date de clôture", self.date_var, 12)
        self.tighten_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="Retirer les lignes vides en série (espaces inutiles)", variable=self.tighten_var).grid(
            row=0, column=2, sticky="w")

        f = ttk.Frame(R)
        f.pack(fill="x", pady=(4, 6))
        self.go_btn = ttk.Button(f, text="Finaliser le Word", command=self._finalize)
        self.go_btn.pack(side="left")
        self.open_btn = ttk.Button(f, text="Ouvrir le Word", command=self._open_word, state="disabled")
        self.open_btn.pack(side="left", padx=8)
        self.progress_var = tk.StringVar()
        ttk.Label(R, textvariable=self.progress_var, foreground=GREY, wraplength=780, justify="left").pack(anchor="w")
        self.final_path = None

    # -- dossier -------------------------------------------------------------
    def _choose_folder(self):
        start = self.cfg.get("last_entree_parent", "") or None
        folder = filedialog.askdirectory(title="Dossier du n° (entrée locative)", initialdir=start)
        if not folder:
            return
        self.cfg["last_entree_parent"] = os.path.dirname(folder)
        self.save_cfg()
        try:
            ee.securite.check_dossier_folder(folder)
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning("Dossier", str(exc))
            return
        if not confirm_own_folder(self, self.cfg, folder):
            return
        self.folder_var.set(folder)
        self._load_folder(folder)

    def _load_folder(self, folder):
        self.folder = folder
        try:
            inv = ee.scan_folder(folder)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Dossier", str(exc))
            return
        self.inv = inv
        names = [os.path.basename(p) for p in inv["docx"]]
        self.docx_combo.configure(values=names)
        if names:
            info = ee.parse_folder_name(folder)
            final = ee.final_docx_name(folder) or ""
            pick = next((n for n in names if n != final), names[0])
            self.docx_var.set(pick)
            self.newname_var.set(final or pick)
        else:
            self.docx_var.set("")
        self.doc_combo.configure(values=[os.path.basename(p) for p in inv["documents"]])
        self.doc_paths = list(inv["documents"])
        if self.doc_paths:
            bail = next((p for p in self.doc_paths if "BAIL" in ee._norm(os.path.basename(p))), self.doc_paths[0])
            self.doc_var.set(os.path.basename(bail))
        self.info_var.set(
            f"{len(names)} Word trouvé(s) • {len(inv['root_images'])} photo(s) en vrac • "
            f"{len(inv['photo_images'])} dans « Photos » • documents : "
            + (", ".join(os.path.basename(p) for p in inv["documents"]) or "aucun bail / OM détecté"))
        if not names:
            messagebox.showwarning("Word introuvable", "Aucun fichier Word (.docx) dans ce dossier.")
        self.photo_items = ee.classify_photos(inv["photo_images"] + inv["root_images"])
        self._update_photo_label()
        self._load_parties()

    def _docx_path(self):
        n = self.docx_var.get()
        return os.path.join(self.folder, n) if self.folder and n else None

    # -- photos --------------------------------------------------------------
    def _ranger(self):
        if not self.folder:
            return
        try:
            moved = ee.ranger_photos(self.folder)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Rangement impossible", str(exc))
            return
        self._load_folder(self.folder)
        self.progress_var.set(f"{len(moved)} photo(s) déplacée(s) dans « Photos ».")

    def _update_photo_label(self):
        n = sum(1 for _p, c, _n in self.photo_items if c)
        notes = [f"{os.path.basename(p)} : {nt}" for p, _c, nt in self.photo_items if nt]
        txt = f"{n} photo(s) retenue(s) sur {len(self.photo_items)}."
        if notes:
            txt += "  À vérifier : " + " ; ".join(notes[:4]) + (" …" if len(notes) > 4 else "")
        txt += "\nPensez à décocher les photos d'index compteurs (bouton « Choisir les photos… »)."
        self.photo_var.set(txt)

    def _pick_photos(self):
        if not self.photo_items:
            messagebox.showinfo("Photos", "Aucune photo dans ce dossier.")
            return

        def done(items):
            self.photo_items = items
            self._update_photo_label()

        PhotoPicker(self, self.photo_items, done)

    # -- parties -------------------------------------------------------------
    def _load_parties(self):
        for w in self.parties_frame.winfo_children():
            w.destroy()
        self.party_widgets = []
        p = self._docx_path()
        if not p:
            return
        try:
            self.parties = ee.read_parties(p)
        except Exception as exc:  # noqa: BLE001
            ttk.Label(self.parties_frame, text=f"Lecture du Word impossible : {exc}", foreground=WARN_ORANGE).pack(anchor="w")
            return
        if not self.parties:
            ttk.Label(self.parties_frame, text="Paragraphes des parties non trouvés (« Locataire, d'une part » / "
                      "« Bailleur …, d'autre part »).", foreground=WARN_ORANGE).pack(anchor="w")
            return
        for part in self.parties:
            row = ttk.Frame(self.parties_frame)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=part["role"], width=10, font=("Segoe UI", 10, "bold")).pack(side="left", anchor="n")
            txt = tk.Text(row, height=2, wrap="word", font=("Segoe UI", 10))
            txt.insert("1.0", part["texte"])
            txt.pack(side="left", fill="x", expand=True)
            right = ttk.Frame(row)
            right.pack(side="left", padx=6, anchor="n")
            hv = tk.BooleanVar(value=False)
            ttk.Checkbutton(right, text="Surligner en jaune", variable=hv).pack(anchor="w")
            st = tk.StringVar()
            ttk.Label(right, textvariable=st, wraplength=220, justify="left").pack(anchor="w")
            self.party_widgets.append((txt, hv, st))
        ttk.Label(self.parties_frame, foreground=GREY, text=(
            "Corrigez directement le texte si vous êtes sûr après vérification ; cochez « Surligner » en cas de doute "
            "ou de différence entre le bail et l'OM.")).pack(anchor="w", pady=(4, 0))

    # -- documents -----------------------------------------------------------
    def _refresh_ocr_status(self):
        exe = ee.find_tesseract()
        if exe:
            self.ocr_var.set("Lecture des scans : Tesseract installé ✓ (noms lus sur les cartes d'identité jointes au bail).")
        else:
            self.ocr_var.set("Lecture des scans : Tesseract non installé. Les PDF « texte » sont lus quand même. Pour lire "
                             "les scans et les cartes d'identité, installez Tesseract (gratuit) : "
                             "github.com/UB-Mannheim/tesseract/wiki — puis relancez le programme.")

    def _current_doc(self):
        n = self.doc_var.get()
        return next((p for p in getattr(self, "doc_paths", []) if os.path.basename(p) == n), None)

    def _add_doc(self):
        ps = filedialog.askopenfilenames(title="Bail, OM, convention…", initialdir=self.folder or None,
                                         filetypes=[("Documents", "*.pdf *.docx *.jpg *.jpeg *.png"), ("Tous", "*.*")])
        if not ps:
            return
        self.doc_paths = list(getattr(self, "doc_paths", [])) + [p for p in ps if p not in getattr(self, "doc_paths", [])]
        self.doc_combo.configure(values=[os.path.basename(p) for p in self.doc_paths])
        self.doc_var.set(os.path.basename(ps[0]))

    def _view_doc(self):
        p = self._current_doc()
        if p:
            DocViewer(self, p)

    def _read_docs(self):
        paths = list(getattr(self, "doc_paths", []))
        if not paths:
            messagebox.showinfo("Documents", "Aucun bail / OM : utilisez « Ajouter un document… ».")
            return

        def progress(msg):
            self._pmsg = msg  # affiché par _tick (jamais d'appel Tk depuis le thread)

        def work():
            return {p: ee.read_document(p, progress_cb=progress) for p in paths}

        def done(results):
            self._pmsg = None
            self.doc_results = results
            self.progress_var.set("Lecture terminée.")
            self._show_doc_results()

        self.progress_var.set("Lecture des documents…")
        run_async(self, work, done, lambda e: messagebox.showerror("Lecture impossible", str(e)))

    def _show_doc_results(self):
        lines, all_names = [], []
        for p, r in self.doc_results.items():
            lines.append(f"■ {os.path.basename(p)} — " + {"texte": "texte lu", "ocr": "lu par OCR", "aucun": "non lu"}[r["source"]])
            if r["remarque"]:
                lines.append("   " + r["remarque"])
            if r["mrz_names"]:
                lines.append("   Noms des cartes d'identité : " + " ; ".join(r["mrz_names"]))
                all_names += [n for n in r["mrz_names"] if n not in all_names]
            if r["source"] in ("texte", "ocr"):
                for page, ctx in ee.find_party_lines(r["text_pages"])[:8]:
                    lines.append(f"   p.{page} : {ctx}")
        self.found.configure(state="normal")
        self.found.delete("1.0", "end")
        self.found.insert("1.0", "\n".join(lines) or "Rien de lisible trouvé.")
        self.found.configure(state="disabled")
        for (txt, hv, st), part in zip(self.party_widgets, self.parties):
            if not all_names:
                st.set("")
                continue
            name, ratio = ee.compare_names(txt.get("1.0", "end"), all_names)
            if ratio >= 0.99:
                st.set(f"✓ identique à la carte d'identité ({name})")
            elif ratio >= 0.6:
                st.set(f"⚠ proche de « {name} » — à vérifier")
                hv.set(True)
            else:
                st.set("Aucun nom correspondant sur les cartes d'identité")

    # -- finalisation ----------------------------------------------------------
    def _finalize(self):
        p = self._docx_path()
        if not p or not os.path.exists(p):
            messagebox.showwarning("Word", "Choisissez le dossier et le Word à finaliser.")
            return
        photos = [ph for ph, c, _n in self.photo_items if c]
        if not photos and not messagebox.askyesno("Aucune photo", "Aucune photo sélectionnée. Continuer quand même ?"):
            return
        # Photos encore en vrac → rangées d'abord dans « Photos »
        if any(os.path.dirname(ph) == self.folder for ph in photos):
            moved = ee.ranger_photos(self.folder)
            mapping = {os.path.join(self.folder, os.path.basename(m)): m for m in moved}
            photos = [mapping.get(ph, ph) for ph in photos]
        updates = []
        for (txt, hv, _st), part in zip(self.party_widgets, self.parties):
            updates.append({"index": part["index"], "texte": txt.get("1.0", "end").strip(), "surligner": hv.get()})
        ensure_signature(self)
        newname = self.newname_var.get().strip()
        if newname and not newname.lower().endswith(".docx"):
            newname += ".docx"
        if not messagebox.askyesno("Confirmer", f"Le Word « {os.path.basename(p)} » va être finalisé et écrasé "
                                   f"({len(photos)} photo(s)), puis renommé en « {newname} ».\n\n"
                                   "Une copie de l'original est d'abord gardée hors Dropbox "
                                   "(Aide → Ouvrir les sauvegardes).\n\nContinuer ?"):
            return
        self.go_btn.configure(state="disabled")

        def progress(msg):
            self._pmsg = msg  # affiché par _tick (jamais d'appel Tk depuis le thread)

        def work():
            return ee.finalize(p, photos, date_cloture=self.date_var.get().strip(), parties_updates=updates,
                               rename_to=newname or None, tighten_layout=self.tighten_var.get(), progress_cb=progress)

        def done(res):
            self._pmsg = None
            self.go_btn.configure(state="normal")
            self.final_path = res["path"]
            self.open_btn.configure(state="normal")
            msg = (f"Word finalisé : {os.path.basename(res['path'])} ({res['photos']} photos).\n"
                   f"Copie de l'original avant modification : {res.get('sauvegarde')}")
            if res["remarques"]:
                msg += "\n\nÀ noter :\n• " + "\n• ".join(res["remarques"])
            self.progress_var.set(msg)
            self._load_folder(self.folder)
            if messagebox.askyesno("Terminé", msg + "\n\nOuvrir le Word pour vérifier ?"):
                open_path(res["path"])

        def failed(exc):
            self._pmsg = None
            self.go_btn.configure(state="normal")
            self.progress_var.set("")
            messagebox.showerror("Finalisation impossible", str(exc))

        run_async(self, work, done, failed)

    def _open_word(self):
        if self.final_path and os.path.exists(self.final_path):
            open_path(self.final_path)


class EntreeBatchWindow(tk.Toplevel):
    """Plusieurs dossiers d'entrée d'un coup : on lance avant de partir,
    les Word sont prêts à la vérification en rentrant."""

    def __init__(self, parent, cfg, save_cfg):
        super().__init__(parent)
        self.title("Entrées — plusieurs dossiers")
        self.geometry("1000x720")
        self.cfg, self.save_cfg = cfg, save_cfg
        self.jobs = []          # dicts de find_pending + "photos"
        self._pmsg = None
        self._stop = False
        self.running = False
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, wraplength=960, justify="left", foreground=GREY, text=(
            "Choisissez le dossier de tranche (ex. « 2026.3100 à 2026.3199 ») : le programme liste VOS entrées (initiales "
            "ci-dessous, ex. GB) dont le Word n'a pas encore de photos ; les dossiers des collègues sont ignorés. "
            "Retirez ceux à ne pas traiter, puis lancez. Chaque dossier est traité l'un après l'autre : "
            "photos rangées et injectées (capture d'écran exclue), parties comparées aux cartes d'identité du bail "
            "(différence → surligné en jaune, le texte n'est jamais réécrit), date, signature, renommage, avec une copie "
            "de chaque original hors Dropbox. Une erreur sur un dossier ne touche ni son Word ni les suivants. "
            "Laissez le PC allumé (la mise en veille est bloquée pendant le lot).")).pack(anchor="w")
        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(8, 4))
        self.parent_var = tk.StringVar(value=cfg.get("last_entree_batch", ""))
        ttk.Label(row, text="Initiales :").pack(side="left")
        self.ini_var = tk.StringVar(value=cfg.get("mes_initiales", "GB"))
        ttk.Entry(row, textvariable=self.ini_var, width=6).pack(side="left", padx=(4, 12))
        ttk.Entry(row, textvariable=self.parent_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Parcourir…", command=self._choose).pack(side="left", padx=6)
        ttk.Button(row, text="Rechercher", command=self._scan).pack(side="left")
        cols = ("dossier", "photos", "statut")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="extended", height=12)
        for c, h, w in (("dossier", "Dossier", 520), ("photos", "Photos", 170), ("statut", "Statut", 250)):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, pady=(4, 4))
        b = ttk.Frame(frm)
        b.pack(fill="x")
        ttk.Label(b, text="Sélection (Ctrl+clic pour plusieurs) :").pack(side="left")
        ttk.Button(b, text="Choisir les photos du dossier sélectionné…", command=self._pick).pack(side="left", padx=6)
        ttk.Button(b, text="Retirer de la liste", command=self._remove).pack(side="left")
        c = ttk.Frame(frm)
        c.pack(fill="x", pady=(8, 0))
        self.date_var = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        ttk.Label(c, text="Date de clôture :").pack(side="left")
        ttk.Entry(c, textvariable=self.date_var, width=12).pack(side="left", padx=(4, 16))
        self.go = ttk.Button(c, text="Lancer le traitement de la liste", command=self._run)
        self.go.pack(side="left")
        self.stop_btn = ttk.Button(c, text="Arrêter après le dossier en cours", command=self._ask_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        self.status = tk.StringVar()
        ttk.Label(frm, textvariable=self.status, foreground=GREY, wraplength=960, justify="left").pack(anchor="w", pady=(6, 0))
        self.report = tk.Text(frm, height=10, wrap="word", font=("Segoe UI", 9))
        self.report.pack(fill="both", expand=True, pady=(6, 0))
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._tick()

    def _tick(self):
        if self._pmsg is not None:
            k, n, name, msg = self._pmsg
            self.status.set(f"Dossier {k}/{n} — {name} : {msg}")
            self._pmsg = None
        try:
            self.after(250, self._tick)
        except tk.TclError:
            pass

    def _choose(self):
        d = filedialog.askdirectory(parent=self, title="Dossier de tranche (ou dossier d'un n°)",
                                    initialdir=self.parent_var.get() or None)
        if d:
            self.parent_var.set(d)
            self._scan()

    def _scan(self):
        parent = self.parent_var.get().strip()
        if not parent or not os.path.isdir(parent):
            return
        ini = self.ini_var.get().strip().upper()
        if not ini:
            messagebox.showwarning("Initiales", "Indiquez vos initiales (ex. GB).", parent=self)
            return
        self.cfg["last_entree_batch"] = parent
        self.cfg["mes_initiales"] = ini
        self.save_cfg()
        self.status.set("Recherche des dossiers à finaliser…")
        skipped = []

        def done(found):
            self.jobs = [dict(j, photos=None) for j in found]
            self._refresh()
            autres = sorted({i for _f, i in skipped})
            self.status.set(f"{len(self.jobs)} dossier(s) d'entrée {ini} à finaliser trouvé(s)."
                            + (f"  Ignorés : {len(skipped)} dossier(s) d'entrée d'autres initiales ({', '.join(autres)})."
                               if skipped else ""))

        run_async(self, lambda: ee.find_pending(parent, initiales=ini, skipped=skipped), done,
                  lambda e: messagebox.showerror("Recherche", str(e), parent=self))

    def _refresh(self, statuses=None):
        self.tree.delete(*self.tree.get_children())
        for i, j in enumerate(self.jobs):
            ph = "automatique" if j["photos"] is None else f"{len(j['photos'])} choisie(s)"
            st = (statuses or {}).get(i, "à traiter")
            self.tree.insert("", "end", iid=str(i), values=(os.path.basename(j["folder"]), ph, st))

    def _pick(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        job = self.jobs[i]
        inv = ee.scan_folder(job["folder"])
        items = ee.classify_photos(inv["photo_images"] + inv["root_images"])
        if job["photos"] is not None:
            chosen = {os.path.basename(p) for p in job["photos"]}
            items = [(p, os.path.basename(p) in chosen, n) for p, _c, n in items]

        def done(sel_items):
            job["photos"] = [p for p, c, _n in sel_items if c]
            self._refresh()

        PhotoPicker(self, items, done)

    def _remove(self):
        for iid in sorted((int(s) for s in self.tree.selection()), reverse=True):
            del self.jobs[iid]
        self._refresh()

    def _ask_stop(self):
        self._stop = True
        self.status.set("Arrêt demandé : le dossier en cours se termine normalement.")

    def _run(self):
        if not self.jobs or self.running:
            return
        if not messagebox.askyesno("Lancer", f"Traiter {len(self.jobs)} dossier(s) maintenant ?\n\n"
                                   "Chaque Word est sauvegardé hors Dropbox avant d'être finalisé.", parent=self):
            return
        ensure_signature(self)
        self.running, self._stop = True, False
        self.go.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        ee.keep_awake(True)
        jobs = [dict(j) for j in self.jobs]
        date_cl = self.date_var.get().strip()

        def progress(k, n, name, msg):
            self._pmsg = (k, n, name, msg)

        def done(results):
            ee.keep_awake(False)
            self.running = False
            self._pmsg = None
            self.go.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            text = ee.batch_report_text(results)
            statuses = {i: ("✓ finalisé" if r["ok"] else "✗ erreur — Word non modifié") for i, r in enumerate(results)}
            self._refresh(statuses)
            self.report.delete("1.0", "end")
            self.report.insert("1.0", text)
            try:
                rp = ee.securite.backup_dir() / f"Rapport lot entrées {date.today():%Y-%m-%d} {datetime_now()}.txt"
                rp.write_text(text, encoding="utf-8")
            except Exception:  # noqa: BLE001
                pass
            self.status.set(f"Lot terminé : {sum(r['ok'] for r in results)}/{len(results)} dossier(s) finalisé(s). "
                            "Vérifiez chaque Word avant envoi.")

        def failed(exc):
            ee.keep_awake(False)
            self.running = False
            self.go.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            messagebox.showerror("Lot interrompu", str(exc), parent=self)

        run_async(self, lambda: ee.process_batch(jobs, date_cloture=date_cl, progress_cb=progress,
                                                 stop_flag=lambda: self._stop), done, failed)

    def _close(self):
        if self.running and not messagebox.askyesno(
                "Lot en cours", "Un lot est en cours. Fermer quand même la fenêtre ?\n"
                "(le dossier en cours se termine, les suivants ne seront pas traités)", parent=self):
            return
        self._stop = True
        self.destroy()


def datetime_now():
    from datetime import datetime

    return datetime.now().strftime("%Hh%M")
