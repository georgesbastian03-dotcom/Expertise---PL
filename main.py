#!/usr/bin/env python3
"""Point d'entrée : lance l'interface graphique."""
import sys
import traceback
from tkinter import messagebox

try:
    import gui
except Exception:  # noqa: BLE001
    # Erreur au chargement (ex. openpyxl absent) : message clair plutôt
    # qu'une fenêtre de console qui se ferme aussitôt.
    err = traceback.format_exc()
    print(err, file=sys.stderr)
    try:
        import tkinter as tk

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Erreur au démarrage",
            "L'application n'a pas pu démarrer.\n\n"
            "Avez-vous installé les dépendances ? Ouvrez une invite de commande "
            "dans ce dossier et lancez :\n\n    pip install -r requirements.txt\n\n"
            f"Détail technique :\n{err[-800:]}",
        )
    except Exception:
        pass
    sys.exit(1)


if __name__ == "__main__":
    gui.main()
