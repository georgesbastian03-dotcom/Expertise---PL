"""Interface graphique (tkinter, inclus avec Python — aucune installation
supplémentaire requise pour l'affichage). Un onglet fonctionnel pour le
permis de location, onglets « État des lieux — entrée » et « — sortie »
(voir gui_etats.py), onglet Expertise réservé.
"""
import os
import shutil
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox

import accueil
import config
import generator
import engine
import history
import plans
import gui_etats
import updater


def _run_async(work_fn, done_fn, widget):
    """Exécute work_fn() dans un thread séparé, puis appelle done_fn(resultat)
    sur le thread principal (via widget.after) une fois terminé.

    Sert à éviter que l'interface se bloque plusieurs minutes, sans aucun
    message, lorsqu'un chemin mémorisé (dossier des gabarits, plans déjà
    sélectionnés) se trouve sur un lecteur réseau ou une clé USB devenue
    inaccessible : Windows peut mettre très longtemps à répondre "introuvable"
    dans ce cas, et cette vérification se faisait jusqu'ici directement au
    démarrage, avant même l'affichage de la fenêtre. Le thread est démon
    (`daemon=True`) pour ne jamais empêcher le programme de se fermer, même
    s'il reste bloqué indéfiniment sur un chemin mort."""

    def worker():
        try:
            result = work_fn()
        except Exception as exc:  # noqa: BLE001
            result = exc
        try:
            widget.after(0, lambda: done_fn(result))
        except Exception:
            pass  # la fenêtre a pu être fermée entre-temps

    threading.Thread(target=worker, daemon=True).start()


ANNEXE_LABELS = {
    "1A": "Annexe 1A — sans règlement communal, logements existants",
    "1B": "Annexe 1B — avec règlement communal, logements existants",
    "1C": "Annexe 1C — sans règlement communal, créés après 01/01/2008",
    "1D": "Annexe 1D — avec règlement communal, créés après 01/01/2008",
}

FORMAT_LABELS = {
    "collectif": "Logement collectif",
    "individuel": "Logement individuel",
}

# Version courte utilisée dans la grille à deux colonnes (type x annexe),
# pour que les deux colonnes tiennent côte à côte sans être coupées.
ANNEXE_LABELS_SHORT = {
    "1A": "1A — sans règlement, existants",
    "1B": "1B — avec règlement, existants",
    "1C": "1C — sans règlement, après 2008",
    "1D": "1D — avec règlement, après 2008",
}


class PermisLocationTab(ttk.Frame):
    def __init__(self, parent, cfg):
        super().__init__(parent, padding=0)
        self.cfg = cfg
        self.mission = None  # (format, units, map, unmatched, ignored, extra_scalars)
        self.excel_path = None
        # Vérifiés en arrière-plan juste après l'affichage (voir
        # _load_last_plan_paths_async) plutôt qu'ici : si un de ces chemins
        # mémorisés est sur un lecteur réseau/une clé USB débranchée, les
        # vérifier tout de suite pourrait bloquer l'ouverture de la fenêtre
        # plusieurs minutes. On capture la liste mémorisée dès maintenant,
        # car _refresh_plans_listbox (appelée par _build_plans_row juste en
        # dessous) réenregistre aussitôt self.plan_paths dans la config — il
        # ne faut pas que cette première sauvegarde (avec une liste encore
        # vide) écrase les chemins mémorisés avant même leur vérification.
        self.plan_paths = []
        self._initial_plan_paths = list(self.cfg.get("last_plan_paths", []))

        # Contenu placé dans un canevas défilant : sur un petit écran ou une
        # fenêtre non maximisée, le bouton "Générer" ne se retrouve jamais
        # caché en bas — on peut toujours faire défiler pour l'atteindre.
        canvas = tk.Canvas(self, highlightthickness=0)
        vscroll = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        self.content = ttk.Frame(canvas, padding=14)
        content_id = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def on_content_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            canvas.itemconfig(content_id, width=event.width)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Le défilement à la molette n'est actif que lorsque la souris
        # survole cet onglet, pour ne pas perturber les autres onglets ou
        # les fenêtres de dialogue (recherche de dossier, mise à jour...).
        def bind_mousewheel(_event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)

        def unbind_mousewheel(_event):
            canvas.unbind_all("<MouseWheel>")

        self.content.bind("<Configure>", on_content_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.bind("<Enter>", bind_mousewheel)
        canvas.bind("<Leave>", unbind_mousewheel)

        self._build_gabarits_row()
        self._build_annexe_row()
        self._build_dossier_row()
        self._build_excel_row()
        self._build_plans_row()
        self._build_preview()
        self._build_generate_row()

        self._refresh_gabarits_status()
        self._load_last_plan_paths_async()

    # ---- Dossier des gabarits ------------------------------------------
    def _build_gabarits_row(self):
        frame = ttk.LabelFrame(self.content, text="Dossier des gabarits", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        self.gabarits_var = tk.StringVar(value=self.cfg.get("gabarits_folder", ""))
        entry = ttk.Entry(frame, textvariable=self.gabarits_var, state="readonly")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(frame, text="Changer...", command=self._choose_gabarits_folder).pack(side="left", padx=(0, 6))
        ttk.Button(frame, text="Ouvrir le dossier", command=self._open_gabarits_folder).pack(side="left")

        self.gabarits_status = ttk.Label(self.content, text="", foreground="#555555")
        self.gabarits_status.pack(fill="x", pady=(0, 10))

    def _choose_gabarits_folder(self):
        folder = filedialog.askdirectory(title="Choisir le dossier contenant les 8 gabarits Word")
        if folder:
            self.gabarits_var.set(folder)
            self.cfg["gabarits_folder"] = folder
            config.save_config(self.cfg)
            self._refresh_gabarits_status()

    def _open_gabarits_folder(self):
        folder = self.gabarits_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showinfo("Dossier des gabarits", "Choisissez d'abord un dossier valide.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # noqa: S606
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le dossier : {exc}")

    def _refresh_gabarits_status(self):
        folder = self.gabarits_var.get()
        if not folder:
            self.gabarits_status.config(text="Aucun dossier sélectionné.", foreground="#a3312f")
            return
        # Vérification faite en arrière-plan (voir _run_async) : si ce dossier
        # est sur un lecteur réseau devenu inaccessible, la vérification peut
        # prendre très longtemps côté Windows ; l'interface ne doit pas se
        # figer en l'attendant.
        self.gabarits_status.config(text="Vérification du dossier…", foreground="#555555")

        def check():
            return generator.check_gabarits_folder(folder)

        def done(result):
            if self.gabarits_var.get() != folder:
                return  # l'utilisateur a changé de dossier entre-temps
            if isinstance(result, Exception):
                self.gabarits_status.config(
                    text=f"Dossier introuvable ou inaccessible : {result}",
                    foreground="#a3312f",
                )
                return
            missing = [s["filename"] for s in result if not s["found"]]
            if not missing:
                self.gabarits_status.config(text="Les 8 gabarits sont présents.", foreground="#2f7a4a")
            else:
                self.gabarits_status.config(
                    text=f"{len(missing)} gabarit(s) manquant(s) : " + ", ".join(missing),
                    foreground="#a3312f",
                )

        _run_async(check, done, self)

    # ---- Type d'annexe (logement collectif / individuel x annexe 1A-1D) --
    def _build_annexe_row(self):
        frame = ttk.LabelFrame(self.content, text="Type d'annexe", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        default_value = self.cfg.get("last_type_annexe")
        if not default_value:
            # Migration depuis l'ancien réglage qui ne mémorisait que l'annexe.
            default_value = "individuel:" + self.cfg.get("last_annexe", "1B")
        self.type_annexe_var = tk.StringVar(value=default_value)

        columns = ttk.Frame(frame)
        columns.pack(fill="x")
        for col, (fmt_code, fmt_label) in enumerate(FORMAT_LABELS.items()):
            col_frame = ttk.Frame(columns)
            col_frame.grid(row=0, column=col, sticky="nw", padx=(0, 30) if col == 0 else 0)
            ttk.Label(col_frame, text=fmt_label, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))
            for code, label in ANNEXE_LABELS_SHORT.items():
                ttk.Radiobutton(
                    col_frame,
                    text=label,
                    value=f"{fmt_code}:{code}",
                    variable=self.type_annexe_var,
                    command=self._on_annexe_change,
                ).pack(anchor="w", pady=(0, 2))

    def _on_annexe_change(self):
        self.cfg["last_type_annexe"] = self.type_annexe_var.get()
        config.save_config(self.cfg)

    def _selected_format_annexe(self):
        """Renvoie (format, annexe) choisis dans la grille, ex: ('individuel', '1B')."""
        fmt, _, annexe = self.type_annexe_var.get().partition(":")
        return fmt, annexe

    # ---- Numéro de dossier (historique / renouvellement à 5 ans) --------
    def _build_dossier_row(self):
        frame = ttk.LabelFrame(self.content, text="Numéro de dossier", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill="x")
        ttk.Label(row, text="N° de dossier :").pack(side="left")
        self.dossier_var = tk.StringVar(value="")
        ttk.Entry(row, textvariable=self.dossier_var, width=20).pack(side="left", padx=(6, 12))
        ttk.Button(row, text="Rechercher un ancien dossier...", command=self._open_dossier_search).pack(side="left")

        ttk.Label(
            frame,
            text=(
                "Facultatif, mais utile pour retrouver cette mission dans 5 ans lors du "
                "renouvellement du permis (le numéro est saisi ici, il n'est pas lu dans l'Excel)."
            ),
            foreground="#5c6570",
            wraplength=640,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

    def _open_dossier_search(self):
        initial = self.dossier_var.get().strip()
        win = tk.Toplevel(self)
        win.title("Rechercher un ancien dossier")
        win.geometry("620x380")
        win.transient(self)

        top = ttk.Frame(win, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="N° de dossier :").pack(side="left")
        search_var = tk.StringVar(value=initial)
        ttk.Entry(top, textvariable=search_var, width=24).pack(side="left", padx=(6, 12))

        result_frame = ttk.Frame(win, padding=(10, 0))
        result_frame.pack(fill="both", expand=True)
        columns = ("date", "annexe", "adresses")
        tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=8)
        tree.heading("date", text="Date de génération")
        tree.heading("annexe", text="Annexe")
        tree.heading("adresses", text="Adresse(s)")
        tree.column("date", width=140)
        tree.column("annexe", width=70, anchor="center")
        tree.column("adresses", width=340)
        tree.pack(fill="both", expand=True, pady=(6, 6))

        matches_holder = {"list": []}

        def do_search():
            matches = history.find_by_dossier(search_var.get())
            matches_holder["list"] = matches
            tree.delete(*tree.get_children())
            for e in matches:
                date_txt = str(e.get("date_generation", ""))[:16].replace("T", " ")
                tree.insert(
                    "", "end",
                    values=(date_txt, e.get("annexe", ""), ", ".join(e.get("adresses") or [])),
                )
            if not matches:
                messagebox.showinfo("Recherche", "Aucun dossier trouvé avec ce numéro.", parent=win)

        ttk.Button(top, text="Rechercher", command=do_search).pack(side="left")

        def get_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo(
                    "Sélection", "Sélectionnez d'abord une ligne dans la liste.", parent=win
                )
                return None
            idx = tree.index(sel[0])
            return matches_holder["list"][idx]

        def open_reports_folder():
            entry = get_selected()
            if not entry:
                return
            folder = entry.get("output_folder")
            if not folder or not os.path.isdir(folder):
                messagebox.showwarning(
                    "Dossier introuvable",
                    "Le dossier des rapports de cette mission n'existe plus à cet emplacement.",
                    parent=win,
                )
                return
            try:
                if sys.platform.startswith("win"):
                    os.startfile(folder)  # noqa: S606
                elif sys.platform == "darwin":
                    os.system(f'open "{folder}"')
                else:
                    os.system(f'xdg-open "{folder}"')
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Erreur", f"Impossible d'ouvrir le dossier : {exc}", parent=win)

        def reuse_mission():
            entry = get_selected()
            if not entry:
                return
            old_excel = entry.get("excel_path")
            if not old_excel or not os.path.isfile(old_excel):
                messagebox.showwarning(
                    "Fichier introuvable",
                    "Le fichier Excel de cette ancienne mission n'existe plus à cet emplacement "
                    "d'origine. Retrouvez-le manuellement puis utilisez Parcourir...",
                    parent=win,
                )
                return

            base, ext = os.path.splitext(old_excel)
            suggested_name = f"{os.path.basename(base)}_{datetime.now().year}{ext}"
            new_path = filedialog.asksaveasfilename(
                title="Enregistrer la copie de l'Excel pour ce renouvellement",
                initialdir=os.path.dirname(old_excel),
                initialfile=suggested_name,
                defaultextension=".xlsx",
                filetypes=[("Fichiers Excel", "*.xlsx")],
                parent=win,
            )
            if not new_path:
                return
            try:
                shutil.copy2(old_excel, new_path)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Erreur", f"Impossible de dupliquer le fichier : {exc}", parent=win)
                return

            # Pré-remplit le numéro de dossier et l'annexe de la mission d'origine.
            self.dossier_var.set(entry.get("dossier", ""))
            if entry.get("annexe") in ANNEXE_LABELS:
                fmt_for_reuse = entry.get("format") if entry.get("format") in FORMAT_LABELS else "individuel"
                self.type_annexe_var.set(f"{fmt_for_reuse}:{entry['annexe']}")
                self._on_annexe_change()

            self.excel_var.set(new_path)
            self.cfg["last_excel_dir"] = os.path.dirname(new_path)
            config.save_config(self.cfg)
            self._load_mission(new_path)

            try:
                if sys.platform.startswith("win"):
                    os.startfile(new_path)  # noqa: S606
            except Exception:
                pass

            messagebox.showinfo(
                "Copie créée",
                "Une copie de l'ancien fichier Excel a été créée et ouverte :\n" + new_path +
                "\n\nMettez à jour la date de passage (et toute autre valeur ayant changé depuis "
                "5 ans), enregistrez dans Excel, puis cliquez sur \"Recharger\" ici avant de "
                "générer les rapports.",
                parent=win,
            )
            win.destroy()

        btns = ttk.Frame(win, padding=10)
        btns.pack(fill="x")
        ttk.Button(btns, text="Ouvrir le dossier des rapports", command=open_reports_folder).pack(side="left")
        ttk.Button(
            btns,
            text="Réutiliser cette mission (dupliquer + mettre à jour)",
            command=reuse_mission,
        ).pack(side="left", padx=(8, 0))

        if initial:
            do_search()

    # ---- Fichier Excel de mission ---------------------------------------
    def _build_excel_row(self):
        frame = ttk.LabelFrame(self.content, text="Fichier Excel de mission", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        self.excel_var = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.excel_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(frame, text="Parcourir...", command=self._choose_excel).pack(side="left", padx=(0, 6))
        ttk.Button(frame, text="Recharger", command=self._reload_excel).pack(side="left")

    def _reload_excel(self):
        if not self.excel_path or not os.path.isfile(self.excel_path):
            messagebox.showinfo("Recharger", "Choisissez d'abord un fichier Excel via Parcourir...")
            return
        self._load_mission(self.excel_path)

    def _choose_excel(self):
        start_dir = self.cfg.get("last_excel_dir") or os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Choisir le fichier Excel de la mission",
            initialdir=start_dir,
            filetypes=[("Fichiers Excel", "*.xlsx")],
        )
        if not path:
            return
        self.excel_var.set(path)
        self.cfg["last_excel_dir"] = os.path.dirname(path)
        config.save_config(self.cfg)
        self._load_mission(path)

    def _load_mission(self, path):
        self.status_var.set(f"Lecture de {os.path.basename(path)}…")
        self.update_idletasks()
        try:
            fmt, units, m, unmatched, ignored, extra_scalars = generator.load_mission(path)
        except generator.MissionError as exc:
            self.mission = None
            self._clear_preview()
            self.status_var.set("")
            messagebox.showerror("Erreur de lecture", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            self.mission = None
            self._clear_preview()
            self.status_var.set("")
            messagebox.showerror("Erreur de lecture", f"Erreur inattendue : {exc}")
            return

        self.excel_path = path
        self.mission = (fmt, units, m, unmatched, ignored, extra_scalars)
        self._render_preview()

    # ---- Plans à joindre ----------------------------------------------------
    def _build_plans_row(self):
        frame = ttk.LabelFrame(self.content, text="Plans à joindre (optionnel)", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            frame,
            text=(
                "Le même jeu de plans (images et/ou PDF) est ajouté en dernières pages de "
                "chaque rapport généré ; le nombre total de pages est indiqué automatiquement "
                "à côté de votre signature (« Nombre d'annexes éventuelles »)."
            ),
            wraplength=640,
            justify="left",
            foreground="#5c6570",
        ).pack(anchor="w", pady=(0, 8))

        body = ttk.Frame(frame)
        body.pack(fill="x")

        list_wrap = ttk.Frame(body)
        list_wrap.pack(side="left", fill="both", expand=True)
        self.plans_listbox = tk.Listbox(list_wrap, height=5, selectmode="extended", exportselection=False)
        self.plans_listbox.pack(side="left", fill="both", expand=True)
        plans_scroll = ttk.Scrollbar(list_wrap, orient="vertical", command=self.plans_listbox.yview)
        plans_scroll.pack(side="left", fill="y")
        self.plans_listbox.configure(yscrollcommand=plans_scroll.set)

        btns = ttk.Frame(body)
        btns.pack(side="left", fill="y", padx=(8, 0))
        ttk.Button(btns, text="Ajouter...", command=self._add_plan_files).pack(fill="x", pady=(0, 4))
        ttk.Button(btns, text="Retirer", command=self._remove_selected_plans).pack(fill="x", pady=(0, 4))
        ttk.Button(btns, text="Monter", command=lambda: self._move_plan(-1)).pack(fill="x", pady=(0, 4))
        ttk.Button(btns, text="Descendre", command=lambda: self._move_plan(1)).pack(fill="x", pady=(0, 4))
        ttk.Button(btns, text="Vider", command=self._clear_plans).pack(fill="x")

        self.plans_count_label = ttk.Label(frame, text="", foreground="#555555")
        self.plans_count_label.pack(fill="x", pady=(6, 0))

        self._refresh_plans_listbox()

    def _refresh_plans_listbox(self):
        self.plans_listbox.delete(0, "end")
        for path in self.plan_paths:
            self.plans_listbox.insert("end", os.path.basename(path))

        if self.plan_paths:
            total = plans.count_total_pages(self.plan_paths)
            self.plans_count_label.config(
                text=f"{len(self.plan_paths)} fichier(s) — {total} page(s) au total seront ajoutées."
            )
        else:
            self.plans_count_label.config(text="Aucun plan sélectionné — aucune page ne sera ajoutée.")

        self.cfg["last_plan_paths"] = list(self.plan_paths)
        config.save_config(self.cfg)

    def _load_last_plan_paths_async(self):
        """Revalide en arrière-plan les plans sélectionnés lors de la
        dernière session (voir le commentaire dans __init__ : un chemin sur
        un lecteur réseau/une clé USB débranchée pourrait sinon bloquer
        l'ouverture de la fenêtre pendant plusieurs minutes)."""
        saved_paths = self._initial_plan_paths
        if not saved_paths:
            return

        def check():
            return [p for p in saved_paths if os.path.isfile(p)]

        def done(result):
            if isinstance(result, Exception) or self.plan_paths:
                return  # échec silencieux, ou l'utilisateur a déjà agi entre-temps
            self.plan_paths = result
            self._refresh_plans_listbox()

        _run_async(check, done, self)

    def _add_plan_files(self):
        start_dir = self.cfg.get("last_plans_dir") or os.path.expanduser("~")
        paths = filedialog.askopenfilenames(
            title="Choisir les pages de plans (images et/ou PDF)",
            initialdir=start_dir,
            filetypes=[
                ("Images et PDF", "*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                ("PDF", "*.pdf"),
                ("Images", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if not paths:
            return
        self.cfg["last_plans_dir"] = os.path.dirname(paths[0])
        config.save_config(self.cfg)
        self.plan_paths.extend(paths)
        self._refresh_plans_listbox()

    def _remove_selected_plans(self):
        selected = set(self.plans_listbox.curselection())
        if not selected:
            return
        self.plan_paths = [p for i, p in enumerate(self.plan_paths) if i not in selected]
        self._refresh_plans_listbox()

    def _move_plan(self, direction):
        selected = list(self.plans_listbox.curselection())
        if len(selected) != 1:
            return
        i = selected[0]
        j = i + direction
        if j < 0 or j >= len(self.plan_paths):
            return
        self.plan_paths[i], self.plan_paths[j] = self.plan_paths[j], self.plan_paths[i]
        self._refresh_plans_listbox()
        self.plans_listbox.selection_set(j)

    def _clear_plans(self):
        if not self.plan_paths:
            return
        self.plan_paths = []
        self._refresh_plans_listbox()

    # ---- Aperçu -----------------------------------------------------------
    def _build_preview(self):
        frame = ttk.LabelFrame(self.content, text="Vérification", padding=10)
        frame.pack(fill="both", expand=True, pady=(0, 10))

        top = ttk.Frame(frame)
        top.pack(fill="x", pady=(0, 6))
        self.format_badge = ttk.Label(top, text="—", font=("Segoe UI", 9, "bold"))
        self.format_badge.pack(side="left")
        self.count_label = ttk.Label(top, text="")
        self.count_label.pack(side="right")

        columns = ("adresse", "n", "bte", "chambres")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=8)
        self.tree.heading("adresse", text="Adresse")
        self.tree.heading("n", text="N°")
        self.tree.heading("bte", text="Bte")
        self.tree.heading("chambres", text="Chambres")
        self.tree.column("adresse", width=260)
        self.tree.column("n", width=60, anchor="e")
        self.tree.column("bte", width=90, anchor="e")
        self.tree.column("chambres", width=80, anchor="e")
        self.tree.pack(fill="both", expand=True)

        self.warning_label = ttk.Label(frame, text="", foreground="#93611c", wraplength=640, justify="left")
        self.warning_label.pack(fill="x", pady=(6, 0))

    def _clear_preview(self):
        self.tree.delete(*self.tree.get_children())
        self.format_badge.config(text="—")
        self.count_label.config(text="")
        self.warning_label.config(text="")
        self.generate_btn.config(state="disabled")

    def _render_preview(self):
        fmt, units, m, unmatched, ignored, extra_scalars = self.mission
        self.tree.delete(*self.tree.get_children())

        self.format_badge.config(text="Logement collectif" if fmt == "collectif" else "Logement individuel")
        self.count_label.config(text=f"{len(units)} logement(s) détecté(s)")

        overflow = []
        for u in units:
            label = engine.unit_label(u, m)
            nb_ch = len(u["chambres"]) if fmt == "collectif" else None
            if nb_ch is not None and nb_ch > engine.MAX_CHAMBRES_TEMPLATE:
                overflow.append(label)
            self.tree.insert("", "end", values=(label["adresse"], label["n"], label["bte"], nb_ch if nb_ch is not None else "—"))

        warnings = []
        if ignored:
            warnings.append(f"{ignored} ligne(s) ignorée(s) (type de logement non reconnu).")
        if overflow:
            warnings.append(
                f"{len(overflow)} logement(s) dépassent les 19 chambres prévues par le gabarit — "
                "les chambres au-delà de la 19e n'apparaîtront pas dans le tableau du rapport."
            )
        self.warning_label.config(text="\n".join(warnings))

        self.generate_btn.config(state="normal")
        self.status_var.set("")

    # ---- Génération ---------------------------------------------------------
    def _build_generate_row(self):
        frame = ttk.Frame(self.content)
        frame.pack(fill="x")

        self.export_pdf_var = tk.BooleanVar(value=self.cfg.get("export_pdf", False))
        ttk.Checkbutton(
            frame,
            text="Exporter aussi en PDF (nécessite Word installé)",
            variable=self.export_pdf_var,
            command=self._on_pdf_toggle,
        ).pack(anchor="w", pady=(0, 8))

        btn_row = ttk.Frame(frame)
        btn_row.pack(anchor="w")
        self.generate_btn = ttk.Button(
            btn_row,
            text="Générer les rapports",
            command=self._on_generate,
            state="disabled",
        )
        self.generate_btn.pack(side="left")
        self.stop_btn = ttk.Button(btn_row, text="Arrêter", command=self._on_stop_generation, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.status_var, foreground="#1f5f6e").pack(anchor="w", pady=(8, 0))

        self.stop_event = None

    def _on_pdf_toggle(self):
        self.cfg["export_pdf"] = self.export_pdf_var.get()
        config.save_config(self.cfg)

    def _on_generate(self):
        if not self.mission:
            return
        gabarits_folder = self.gabarits_var.get()
        if not gabarits_folder or not os.path.isdir(gabarits_folder):
            messagebox.showwarning("Dossier des gabarits", "Choisissez d'abord le dossier contenant vos gabarits Word.")
            return

        selected_format, annexe = self._selected_format_annexe()
        fmt = self.mission[0]
        if selected_format != fmt:
            messagebox.showerror(
                "Type de logement",
                "Vous avez sélectionné « " + FORMAT_LABELS.get(selected_format, selected_format) + " », mais ce "
                "fichier Excel correspond à « " + FORMAT_LABELS.get(fmt, fmt) + " » (détecté automatiquement à "
                "partir de la colonne « Type Logement »).\n\n"
                "Choisissez la bonne colonne (Logement collectif ou Logement individuel) pour cette annexe avant "
                "de générer.",
            )
            return

        status = generator.check_gabarits_folder(gabarits_folder)
        missing_needed = [
            s["filename"] for s in status if s["annexe"] == annexe and s["format"] == fmt and not s["found"]
        ]
        if missing_needed:
            messagebox.showerror(
                "Gabarit manquant",
                "Le gabarit requis est introuvable dans le dossier sélectionné :\n" + ", ".join(missing_needed),
            )
            return

        # On propose toujours un dossier de destination tout prêt (créé s'il
        # n'existe pas encore) à côté de l'Excel, pour qu'il soit déjà
        # sélectionné à l'ouverture de la fenêtre — l'utilisateur peut bien
        # sûr en choisir un autre.
        default_output = os.path.join(os.path.dirname(self.excel_path), "Rapports_Generes")
        try:
            os.makedirs(default_output, exist_ok=True)
        except Exception:
            pass
        output_folder = filedialog.askdirectory(
            title="Dossier de destination des rapports (proposé : Rapports_Generes)",
            initialdir=default_output,
        )
        if not output_folder:
            return

        dossier = self.dossier_var.get().strip()
        self.generate_btn.config(state="disabled")
        self.stop_event = threading.Event()
        self.stop_btn.config(state="normal")
        export_pdf = self.export_pdf_var.get()
        _, units, m, _, _, extra_scalars = self.mission
        total_units = len(units)
        adresses = sorted({engine.unit_label(u, m)["adresse"] for u in units})
        plan_paths = list(self.plan_paths)
        stop_event = self.stop_event

        def worker():
            def progress(i, total, label):
                # Appelé depuis le thread de génération : ne jamais toucher
                # les widgets Tkinter directement, on passe par after().
                text = f"Génération {i} / {total} — {label['adresse']} n°{label['n']}…"
                self.after(0, lambda: self.status_var.set(text))

            if plan_paths:
                self.after(0, lambda: self.status_var.set("Préparation des pages de plans…"))

            try:
                results = generator.generate_reports(
                    annexe, units, m, gabarits_folder, output_folder, export_pdf, progress, extra_scalars,
                    plan_paths, should_stop=stop_event.is_set,
                )
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda: self._on_generation_failed(exc))
                return
            stopped = stop_event.is_set()
            self.after(
                0,
                lambda: self._on_generation_done(results, output_folder, dossier, annexe, fmt, adresses, total_units, stopped),
            )

        threading.Thread(target=worker, daemon=True).start()

    def _on_stop_generation(self):
        if self.stop_event:
            self.stop_event.set()
        self.stop_btn.config(state="disabled")
        self.status_var.set("Arrêt demandé… le logement en cours de génération sera terminé, puis la génération s'arrêtera.")

    def _on_generation_failed(self, exc):
        self.generate_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.status_var.set("")
        messagebox.showerror("Erreur de génération", f"Erreur lors de la préparation des plans ou de la génération :\n{exc}")

    def _on_generation_done(self, results, output_folder, dossier, annexe, fmt, adresses, total_units=None, stopped=False):
        self.generate_btn.config(state="normal")
        self.stop_btn.config(state="disabled")

        # Historique (uniquement si un n° de dossier a été saisi) : permet de
        # retrouver cette mission dans 5 ans lors du renouvellement.
        history.add_entry(dossier, annexe, fmt, self.excel_path, output_folder, adresses)
        errors = [r for r in results if r["error"]]
        ok = [r for r in results if not r["error"]]
        unknown_tags = set()
        for r in results:
            unknown_tags.update(r.get("unknown_tags") or [])

        if stopped:
            self.status_var.set(f"Génération arrêtée : {len(ok)} rapport(s) généré(s) dans {output_folder}.")
        else:
            self.status_var.set(f"{len(ok)} rapport(s) généré(s) dans {output_folder}.")

        total_attendu = total_units if total_units is not None else len(results)
        if stopped:
            msg_lines = [f"Génération arrêtée par vous : {len(ok)} rapport(s) généré(s) sur {total_attendu} avant l'arrêt."]
        else:
            msg_lines = [f"{len(ok)} rapport(s) généré(s) sur {total_attendu}."]
        if unknown_tags:
            msg_lines.append("Balises non reconnues dans le gabarit (laissées vides) : " + ", ".join(sorted(unknown_tags)))
        if errors:
            msg_lines.append("\nErreurs :")
            for e in errors[:10]:
                msg_lines.append(f"- {e['file_base']} : {e['error']}")
            if len(errors) > 10:
                msg_lines.append(f"... et {len(errors) - 10} autre(s).")

        if errors:
            messagebox.showwarning("Génération terminée avec erreurs", "\n".join(msg_lines))
        elif stopped:
            messagebox.showinfo("Génération arrêtée", "\n".join(msg_lines))
        else:
            messagebox.showinfo("Génération terminée", "\n".join(msg_lines))

        try:
            if sys.platform.startswith("win"):
                os.startfile(output_folder)  # noqa: S606
        except Exception:
            pass


class PlaceholderTab(ttk.Frame):
    def __init__(self, parent, title, note):
        super().__init__(parent, padding=24)
        ttk.Label(self, text=title, font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))
        ttk.Label(self, text=note, wraplength=560, justify="left", foreground="#5c6570").pack(anchor="w")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rapports de Visite — Experts Nicolai")
        self.geometry("1040x760")
        self.minsize(660, 480)
        try:
            if sys.platform.startswith("win"):
                self.state("zoomed")  # ouvre en plein écran (barre des tâches conservée)
        except Exception:
            pass

        try:
            style = ttk.Style(self)
            if "clam" in style.theme_names():
                style.theme_use("clam")
            elif "vista" in style.theme_names():
                style.theme_use("vista")
        except Exception:
            pass

        self.cfg = config.load_config()

        self._build_menu()

        self._build_pages()

        # Vérification silencieuse des mises à jour au démarrage : ne dérange
        # l'utilisateur que si une version plus récente existe réellement.
        self.after(800, lambda: self._check_update(silent=True))

    def _build_pages(self):
        """Accueil (Expertise immobilière / Topographie), puis les bulles de
        l'expertise qui ouvrent chacune une page de travail."""
        nav = accueil.Navigateur(self)
        nav.pack(fill="both", expand=True)
        aller = lambda nom: (lambda: nav.afficher(nom))  # noqa: E731
        save_cfg = lambda: config.save_config(self.cfg)  # noqa: E731

        nav.ajouter("accueil", accueil.PageBulles(
            nav, "Experts Nicolai", "Choisissez une section", [
                dict(titre="Expertise immobilière", dessin=accueil.dessin_expertise_immobiliere,
                     commande=aller("expertise"), couleur="#dcebf9"),
                dict(titre="Topographie", dessin=accueil.dessin_topographie,
                     commande=aller("topographie"), couleur="#e3f2e1", sous_titre="En préparation"),
            ], diametre=260))

        nav.ajouter("expertise", accueil.PageModule(
            nav, "Expertise immobilière", aller("accueil"),
            lambda parent: accueil.PageBulles(parent, "Expertise immobilière", "Que voulez-vous encoder ?", [
                dict(titre="Entrées locatives", dessin=accueil.dessin_entree,
                     commande=aller("entree"), couleur="#e3f2e1"),
                dict(titre="Sorties locatives", dessin=accueil.dessin_sortie,
                     commande=aller("sortie"), couleur="#fdebd8"),
                dict(titre="Expertises", dessin=accueil.dessin_expertises,
                     commande=aller("expertises"), couleur="#dcebf9"),
                dict(titre="Permis de location", dessin=accueil.dessin_permis,
                     commande=aller("permis"), couleur="#efe5f7"),
            ], diametre=180)))

        retour = aller("expertise")
        nav.ajouter("entree", accueil.PageModule(
            nav, "État des lieux — entrée", retour,
            lambda parent: gui_etats.EntreeTab(parent, self.cfg, save_cfg)))
        nav.ajouter("sortie", accueil.PageModule(
            nav, "État des lieux — sortie", retour,
            lambda parent: gui_etats.SortieTab(parent, self.cfg, save_cfg)))
        nav.ajouter("expertises", accueil.PageModule(
            nav, "Expertises", retour,
            lambda parent: PlaceholderTab(
                parent, "Expertise",
                "Cette page sera construite dans un second temps, sur un principe "
                "différent de celui du permis de location.")))
        nav.ajouter("permis", accueil.PageModule(
            nav, "Permis de location", retour,
            lambda parent: PermisLocationTab(parent, self.cfg)))
        nav.ajouter("topographie", accueil.PageModule(
            nav, "Topographie", aller("accueil"),
            lambda parent: PlaceholderTab(
                parent, "Topographie",
                "Cette section est en préparation : elle accueillera plus tard "
                "les outils de topographie.")))

        nav.afficher("accueil")

    def _build_menu(self):
        menubar = tk.Menu(self)
        aide_menu = tk.Menu(menubar, tearoff=0)
        aide_menu.add_command(
            label="Vérifier les mises à jour…",
            command=lambda: self._check_update(silent=False),
        )
        aide_menu.add_command(label="Ma signature…", command=lambda: gui_etats.choose_signature(self))
        aide_menu.add_command(label="Ouvrir les sauvegardes",
                              command=lambda: gui_etats.open_path(str(gui_etats.ee.securite.backup_dir())))
        aide_menu.add_separator()
        aide_menu.add_command(label="À propos", command=self._show_about)
        menubar.add_cascade(label="Aide", menu=aide_menu)
        self.config(menu=menubar)

    def _show_about(self):
        messagebox.showinfo(
            "À propos",
            "Rapports de Visite — Experts Nicolai\n"
            f"Version {updater.APP_VERSION}",
        )

    def _check_update(self, silent):
        """Lance la vérification en arrière-plan pour ne jamais geler la
        fenêtre (dépend du réseau)."""

        def worker():
            result = updater.check_for_update()
            self.after(0, lambda: self._on_update_checked(result, silent))

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_checked(self, result, silent):
        if result is None:
            if not silent:
                messagebox.showinfo(
                    "Mises à jour",
                    f"Vous avez déjà la dernière version ({updater.APP_VERSION}).",
                )
            return

        notes = result["notes"]
        detail = f"\n\nDétails :\n{notes}" if notes else ""
        proceed = messagebox.askyesno(
            "Mise à jour disponible",
            f"Une nouvelle version ({result['tag']}) est disponible "
            f"(version actuelle : {updater.APP_VERSION}).{detail}\n\n"
            "Voulez-vous la télécharger et l'installer maintenant ? "
            "Le programme redémarrera automatiquement. Votre dossier de "
            "gabarits et vos réglages ne sont pas affectés.",
        )
        if not proceed:
            return
        self._install_update(result["url"])

    def _install_update(self, zip_url):
        progress = tk.Toplevel(self)
        progress.title("Mise à jour en cours")
        progress.resizable(False, False)
        progress.transient(self)
        progress.grab_set()
        status_var = tk.StringVar(value="Préparation…")
        ttk.Label(progress, textvariable=status_var, padding=20).pack()
        progress.update_idletasks()

        def worker():
            try:
                updater.download_and_install(
                    zip_url,
                    progress_cb=lambda msg: self.after(0, lambda: status_var.set(msg)),
                )
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda: self._on_update_failed(progress, exc))
                return
            self.after(0, updater.relaunch_and_exit)

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_failed(self, progress, exc):
        progress.destroy()
        messagebox.showerror(
            "Échec de la mise à jour",
            "La mise à jour n'a pas pu être installée "
            f"(le programme actuel n'a pas été modifié) :\n\n{exc}",
        )


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
