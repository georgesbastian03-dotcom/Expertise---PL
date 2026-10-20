#!/usr/bin/env python3
"""Point d'entrée : lance l'interface graphique."""
import os
import sys
import traceback

ERREUR_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "erreur.log")


def signaler_erreur(intro):
    """Affiche l'erreur dans une fenêtre et l'écrit dans erreur.log : lancé
    depuis l'icône du Bureau, le programme n'a pas de console visible, donc
    sans cela une erreur ferait simplement « rien ne se passe »."""
    err = traceback.format_exc()
    print(err, file=sys.stderr)
    try:
        with open(ERREUR_LOG, "w", encoding="utf-8") as f:
            f.write(err)
    except Exception:
        pass
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Erreur au démarrage",
            f"{intro}\n\nLe détail a été enregistré dans :\n{ERREUR_LOG}\n\n"
            f"Détail technique :\n{err[-800:]}",
        )
        root.destroy()
    except Exception:
        pass


try:
    import gui
except Exception:  # noqa: BLE001
    # Erreur au chargement (ex. openpyxl absent).
    signaler_erreur(
        "L'application n'a pas pu démarrer.\n\n"
        "Avez-vous installé les dépendances ? Ouvrez une invite de commande "
        "dans ce dossier et lancez :\n\n    pip install -r requirements.txt"
    )
    sys.exit(1)


if __name__ == "__main__":
    try:
        gui.main()
    except Exception:  # noqa: BLE001
        signaler_erreur("L'application s'est arrêtée à cause d'une erreur.")
        sys.exit(1)
