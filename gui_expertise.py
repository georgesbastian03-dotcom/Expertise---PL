"""Onglet « Expertises » : ABEX depuis le modèle vierge et photos du rapport.

Tout fonctionne en local, sans IA, sans internet. L'expert garde la main sur
la lecture des notes, l'analyse et les valeurs : le programme fait la partie
mécanique (ABEX, photos, nommage) sans jamais écraser un fichier existant.
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import expertise_engine as xe
from gui_etats import (GREY, RolePhotoPicker, ScrollArea, confirm_own_folder, open_path, run_async,
                       section)


class ExpertiseTab(ttk.Frame):
    def __init__(self, parent, cfg, save_cfg):
        super().__init__(parent)
        self.cfg, self.save_cfg = cfg, save_cfg
        self.folder = None
        self.photo_items = []
        self.vars = {}
        self.surf_vars = {}
        self._pmsg = None
        area = ScrollArea(self)
        area.pack(fill="both", expand=True)
        self.R = area.body
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
        ttk.Label(R, wraplength=800, justify="left", foreground=GREY, text=(
            "Encodage mécanique d'une expertise : l'ABEX est créé à partir du modèle VIERGE officiel (seul l'onglet "
            "Données est rempli, les formules restent intactes) et les photos sont placées dans le rapport Word "
            "(façade en page de garde, planches 3 × 3 sous « PHOTO(S) DE L'IMMEUBLE »). Aucun fichier existant "
            "n'est écrasé : une nouvelle version « v2 » est écrite à côté. Ne reprenez que ce qui est entouré, coché "
            "ou écrit dans les notes ; laissez vide ce qui n'y figure pas.")).pack(anchor="w", pady=(0, 10))

        f = section(R, "1. Dossier")
        f.columnconfigure(1, weight=1)
        self.folder_var = tk.StringVar()
        ttk.Label(f, text="Dossier du n° :").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.folder_var).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(f, text="Parcourir…", command=self._choose_folder).grid(row=0, column=2)
        self.info_var = tk.StringVar()
        ttk.Label(f, textvariable=self.info_var, foreground=GREY, wraplength=760, justify="left").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        f = section(R, "2. Modèle ABEX vierge")
        f.columnconfigure(1, weight=1)
        self.model_var = tk.StringVar(value=self.cfg.get("modele_abex", ""))
        ttk.Label(f, text="Fichier :").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.model_var, state="readonly").grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(f, text="Changer…", command=self._choose_model).grid(row=0, column=2)
        ttk.Label(f, foreground=GREY, text=(f"« {xe.MODELE_ABEX_NOM} » (dossier « IA MODELES Officiels »). "
                                            "Jamais l'ABEX d'un autre dossier.")).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        f = section(R, "3. Données (onglet Données de l'ABEX)")
        for i, (key, cell, label, _kind) in enumerate(xe.CHAMPS):
            r, c = divmod(i, 2)
            ttk.Label(f, text=f"{label} [{cell}]").grid(row=r, column=c * 2, sticky="w", padx=(0, 6), pady=2)
            v = tk.StringVar()
            ttk.Entry(f, textvariable=v, width=30).grid(row=r, column=c * 2 + 1, sticky="we", padx=(0, 14), pady=2)
            self.vars[key] = v

        f = section(R, "3 bis. Référence cadastrale — le cadastre officiel fait foi (CADGIS > CADMAP)")
        ttk.Label(f, foreground=GREY, wraplength=760, justify="left", text=(
            "Copiez l'encadré « Référence cadastrale » de WalOnMap (ou la fiche CADGIS) puis « Coller l'encadré » : "
            "la CaPaKey, ses parties et la superficie sont remplies (la superficie va en G42). Les notes, l'OM et le "
            "client ne servent qu'à comparer : un écart est signalé, le cadastre est retenu.")).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))
        champs = xe.CADASTRE + xe.NOTES_CAD
        for i, (key, label) in enumerate(champs):
            r, c = divmod(i, 2)
            ttk.Label(f, text=label).grid(row=r + 1, column=c * 2, sticky="w", padx=(0, 6), pady=2)
            v = tk.StringVar()
            e = ttk.Entry(f, textvariable=v, width=30)
            e.grid(row=r + 1, column=c * 2 + 1, sticky="we", padx=(0, 14), pady=2)
            self.vars[key] = v
            if key == "capakey":
                e.bind("<FocusOut>", lambda _e: self._split_capakey())
        r = (len(champs) + 1) // 2 + 1
        b = ttk.Frame(f)
        b.grid(row=r, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Button(b, text="Coller l'encadré WalOnMap / CADGIS", command=self._paste_cadastre).pack(side="left")
        ttk.Button(b, text="Ouvrir une capture de l'encadré…", command=self._open_capture).pack(side="left", padx=6)
        self.cad_var = tk.StringVar()
        ttk.Label(f, textvariable=self.cad_var, foreground="#b45309", wraplength=760, justify="left").grid(
            row=r + 1, column=0, columnspan=4, sticky="w", pady=(4, 0))

        f = section(R, "4. Surfaces brutes (SEM) par niveau, en m² — formule possible : =9*6,98+1,23*2,5")
        for i, (cell, label) in enumerate(xe.SURFACES):
            r, c = divmod(i, 3)
            ttk.Label(f, text=f"{label} [{cell}]").grid(row=r, column=c * 2, sticky="w", padx=(0, 6), pady=2)
            v = tk.StringVar()
            ttk.Entry(f, textvariable=v, width=12).grid(row=r, column=c * 2 + 1, sticky="w", padx=(0, 14), pady=2)
            self.surf_vars[cell] = v
        ttk.Label(f, foreground=GREY, wraplength=760, justify="left", text=(
            "Maison : SEM calculée depuis les cotes extérieures du croquis, sans coefficient. Appartement : "
            "SIM × 1,15 avec la formule visible (ex. =96,22*1,15).")).grid(row=7, column=0, columnspan=6, sticky="w")

        g = ttk.Frame(R)
        g.pack(fill="x", pady=(0, 10))
        self.abex_btn = ttk.Button(g, text="Créer l'ABEX", command=self._create_abex)
        self.abex_btn.pack(side="left")
        ttk.Button(g, text="Enregistrer la saisie", command=self._save).pack(side="left", padx=6)

        f = section(R, "5. Photos du rapport Word")
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Rapport Word :").grid(row=0, column=0, sticky="w")
        self.word_var = tk.StringVar()
        self.word_combo = ttk.Combobox(f, textvariable=self.word_var, state="readonly")
        self.word_combo.grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(f, text="Autre Word…", command=self._choose_word).grid(row=0, column=2)
        self.photo_var = tk.StringVar(value="—")
        ttk.Label(f, textvariable=self.photo_var, wraplength=600, justify="left").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(f, text="Choisir les photos…", command=self._pick_photos).grid(row=1, column=2, pady=(6, 0))
        self.photos_btn = ttk.Button(f, text="Placer les photos dans le Word (nouvelle version)",
                                     command=self._insert_photos)
        self.photos_btn.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.progress_var = tk.StringVar()
        ttk.Label(R, textvariable=self.progress_var, foreground=GREY, wraplength=800, justify="left").pack(anchor="w")

    # -- dossier ---------------------------------------------------------------
    def _choose_folder(self):
        start = self.cfg.get("last_expertise_parent", "") or None
        folder = filedialog.askdirectory(title="Dossier du n° (expertise)", initialdir=start)
        if not folder:
            return
        self.cfg["last_expertise_parent"] = os.path.dirname(folder)
        self.save_cfg()
        try:
            xe.securite.check_dossier_folder(folder)
        except Exception as exc:  # noqa: BLE001
            messagebox.showwarning("Dossier", str(exc))
            return
        if not confirm_own_folder(self, self.cfg, folder):
            return
        self.folder_var.set(folder)
        self._load_folder(folder)

    def _load_folder(self, folder):
        self.folder = folder
        info = xe.parse_folder(folder)
        names = xe.output_names(folder, self.cfg.get("mes_initiales") or "GB")
        data = xe.load_saisie(info["dossier"]) or xe.default_data(folder)
        for key, v in self.vars.items():
            v.set(data.get(key, "") or "")
        for cell, v in self.surf_vars.items():
            v.set((data.get("surfaces") or {}).get(cell, ""))
        self._check_cadastre()
        words = xe.find_words(folder)
        self.word_paths = words
        self.word_combo.configure(values=[os.path.basename(p) for p in words])
        if words:
            self.word_combo.current(len(words) - 1)
        else:
            self.word_var.set("")
        try:
            self.photo_items = xe.photo_items(folder)
        except Exception as exc:  # noqa: BLE001
            self.photo_items = []
            self.photo_var.set(f"Photos illisibles : {exc}")
        self._update_photo_label()
        self.info_var.set(f"Dossier {info['dossier']} — ABEX : « {names['abex']} » • {len(words)} rapport(s) Word "
                          f"trouvé(s). Vérifiez la rue et la localité proposées (tirées du nom du dossier).")

    def _choose_model(self):
        p = filedialog.askopenfilename(title="Modèle ABEX vierge", filetypes=[("Excel", "*.xlsx")],
                                       initialdir=os.path.dirname(self.model_var.get()) or None)
        if p:
            self.cfg["modele_abex"] = p
            self.save_cfg()
            self.model_var.set(p)

    def _choose_word(self):
        if not self.folder:
            return
        p = filedialog.askopenfilename(title="Rapport Word", initialdir=self.folder, filetypes=[("Word", "*.docx")])
        if p:
            self.word_paths = list(getattr(self, "word_paths", [])) + [p]
            self.word_combo.configure(values=[os.path.basename(x) for x in self.word_paths])
            self.word_combo.current(len(self.word_paths) - 1)

    # -- cadastre --------------------------------------------------------------
    def _split_capakey(self):
        parts = xe.parse_capakey(self.vars["capakey"].get())
        if parts:
            for k, v in parts.items():
                if k in self.vars:
                    self.vars[k].set(v)
        self._check_cadastre()

    def _paste_cadastre(self):
        try:
            txt = self.clipboard_get()
        except tk.TclError:
            txt = ""
        found = xe.parse_walonmap(txt)
        if not found:
            messagebox.showinfo("Cadastre", "Rien de reconnu dans le presse-papiers. Copiez le texte de l'encadré "
                                            "« Référence cadastrale » (ou encodez les champs à la main).")
            return
        for k, v in found.items():
            if k in self.vars:
                self.vars[k].set(v)
        if not self.vars["source_cad"].get().strip():
            self.vars["source_cad"].set("WalOnMap")
        self._check_cadastre()

    def _open_capture(self):
        p = filedialog.askopenfilename(title="Capture de l'encadré WalOnMap", initialdir=self.folder or None,
                                       filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp"), ("Tous", "*.*")])
        if p:
            open_path(p)
            self.cad_var.set("Recopiez la CaPaKey et la superficie de la capture dans les champs ci-dessus.")

    def _check_cadastre(self):
        ecarts = xe.ecarts_cadastre({k: v.get() for k, v in self.vars.items()})
        self.cad_var.set(("À signaler : " + " ".join(ecarts)) if ecarts else "")

    # -- données ---------------------------------------------------------------
    def _data(self):
        info = xe.parse_folder(self.folder)
        data = {k: v.get() for k, v in self.vars.items()}
        data["dossier"] = info["dossier"]
        data["surfaces"] = {c: v.get() for c, v in self.surf_vars.items() if v.get().strip()}
        return data

    def _save(self):
        if not self.folder:
            return
        xe.save_saisie(self._data())
        self.progress_var.set("Saisie enregistrée sur ce PC (elle sera reprise à la prochaine ouverture du dossier).")

    def _create_abex(self):
        if not self.folder:
            messagebox.showwarning("Dossier", "Choisissez d'abord le dossier.")
            return
        model = self.model_var.get()
        if not model or not os.path.exists(model):
            messagebox.showwarning("Modèle", f"Choisissez le modèle vierge « {xe.MODELE_ABEX_NOM} ».")
            return
        data = self._data()
        try:
            xe.cell_values(data)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Données", str(exc))
            return
        try:
            res = xe.create_abex(self.folder, data, model, self.cfg.get("mes_initiales") or "GB")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("ABEX", str(exc))
            return
        msg = f"ABEX créé : {os.path.basename(res['path'])}"
        if res["remarques"]:
            msg += "\n\nÀ noter :\n• " + "\n• ".join(res["remarques"])
        self.progress_var.set(msg)
        if messagebox.askyesno("ABEX", msg + "\n\nOuvrir l'ABEX ?"):
            open_path(res["path"])

    # -- photos ------------------------------------------------------------------
    def _update_photo_label(self):
        n = {r: sum(1 for it in self.photo_items if it["role"] == r) for r in xe.ROLES}
        txt = (f"Façade : {n['facade']} • Planches : {n['planche']} photo(s) "
               f"({(n['planche'] + 8) // 9} planche(s) 3 × 3) • Exclues : {n['exclure']}")
        notes = [f"photo {i + 1} : {it['note']}" for i, it in enumerate(self.photo_items) if it.get("note")]
        if notes:
            txt += "\nÀ vérifier : " + " ; ".join(notes[:3])
        self.photo_var.set(txt if self.photo_items else "Aucune photo dans « Photos ».")

    def _pick_photos(self):
        if not self.photo_items:
            messagebox.showinfo("Photos", "Aucune photo dans ce dossier.")
            return

        def done(items):
            self.photo_items = items
            self._update_photo_label()

        RolePhotoPicker(self, self.photo_items, done, roles=xe.ROLES, labels=xe.ROLE_LABELS, consigne=(
            "« Façade (garde) » : la photo de la page de garde (une seule). « Planche » : photos des planches 3 × 3, "
            "dans l'ordre (façade, terrasse, séjour, cuisine, hall, sanitaires, chambres, technique, cave, garage). "
            "Excluez la capture du calendrier. ↻ tourne la photo d'un quart de tour."))

    def _insert_photos(self):
        idx = self.word_combo.current()
        if not self.folder or idx < 0:
            messagebox.showwarning("Word", "Choisissez le dossier et le rapport Word.")
            return
        word = self.word_paths[idx]
        facades = [it for it in self.photo_items if it["role"] == "facade"]
        planches = [it for it in self.photo_items if it["role"] == "planche"]
        if len(facades) > 1:
            messagebox.showwarning("Photos", "Une seule photo « Façade (garde) » est possible.")
            return
        if not facades and not planches:
            messagebox.showwarning("Photos", "Aucune photo retenue.")
            return
        if not messagebox.askyesno("Confirmer", f"Une nouvelle version de « {os.path.basename(word)} » va être écrite "
                                   f"à côté (façade : {'oui' if facades else 'non'}, {len(planches)} photo(s) en "
                                   "planches). L'original n'est pas modifié.\n\nContinuer ?"):
            return
        self.photos_btn.configure(state="disabled")

        def progress(msg):
            self._pmsg = msg

        def done(res):
            self._pmsg = None
            self.photos_btn.configure(state="normal")
            msg = f"Word écrit : {os.path.basename(res['path'])} ({res['planches']} planche(s))."
            if res["remarques"]:
                msg += "\n• " + "\n• ".join(res["remarques"])
            self.progress_var.set(msg)
            self._load_folder(self.folder)
            if messagebox.askyesno("Terminé", msg + "\n\nOuvrir le Word ?"):
                open_path(res["path"])

        def failed(exc):
            self._pmsg = None
            self.photos_btn.configure(state="normal")
            messagebox.showerror("Photos", str(exc))

        run_async(self, lambda: xe.insert_photos(word, facades[0] if facades else None, planches, progress), done,
                  failed)
