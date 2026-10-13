"""Historique local des missions générées, pour retrouver un dossier lors du
renouvellement du permis de location (obligatoire tous les 5 ans) à partir de
son numéro de dossier — saisi manuellement, aucune colonne dédiée n'existant
dans les fichiers Excel de mission.

Stocké à part du dossier du programme (comme config.py) : ni les mises à
jour automatiques ni un remplacement du dossier de l'application n'affectent
cet historique.
"""
import json
import os
from datetime import datetime

import config as config_module


def _history_path():
    return os.path.join(str(config_module._config_dir()), "historique.json")


def load_history():
    path = _history_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def save_history(entries):
    path = _history_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(entries, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # l'historique n'est jamais bloquant pour la génération


def add_entry(dossier, annexe, fmt, excel_path, output_folder, adresses):
    """Enregistre une mission générée. N'enregistre rien si aucun numéro de
    dossier n'a été saisi (mission non retrouvable plus tard, par choix de
    l'utilisateur)."""
    dossier = (dossier or "").strip()
    if not dossier:
        return
    entries = load_history()
    entries.append(
        {
            "dossier": dossier,
            "date_generation": datetime.now().isoformat(timespec="seconds"),
            "annexe": annexe,
            "format": fmt,
            "excel_path": excel_path,
            "output_folder": output_folder,
            "adresses": adresses,
        }
    )
    save_history(entries)


def find_by_dossier(dossier):
    """Renvoie les missions passées pour ce numéro de dossier, les plus
    récentes en premier."""
    dossier = (dossier or "").strip().lower()
    if not dossier:
        return []
    matches = [
        e for e in load_history() if str(e.get("dossier", "")).strip().lower() == dossier
    ]
    matches.sort(key=lambda e: e.get("date_generation", ""), reverse=True)
    return matches
