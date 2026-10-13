"""Garde-fous pour travailler dans la Dropbox du bureau sans jamais rien perdre.

Principes :
  * le programme ne SUPPRIME jamais aucun fichier ;
  * avant d'écraser un fichier existant (Word d'entrée, Excel/PV de sortie),
    une copie de sauvegarde est faite HORS de la Dropbox, dans
    %APPDATA%\\RapportsPermisLocation\\sauvegardes\\ (pas de fichier en trop
    dans les dossiers du bureau) ;
  * les fichiers sont d'abord écrits sous un nom temporaire, puis mis à leur
    place en une seule opération : si le PC plante en cours d'écriture,
    l'original reste intact ;
  * les actions ne s'appliquent qu'à un dossier de dossier « 2026.xxxx … »,
    jamais à la racine de la Dropbox ou à un dossier de tranche.
"""
import os
import re
import shutil
import time
from datetime import datetime

import config

DOSSIER_RE = re.compile(r"^\d{4}\.\d{3,4}\s+\S+")
KEEP_DAYS = 60


class SecuriteError(Exception):
    pass


def backup_dir():
    d = config._config_dir() / "sauvegardes"
    d.mkdir(parents=True, exist_ok=True)
    return d


def check_dossier_folder(folder):
    """Refuse tout dossier qui n'est pas un dossier de dossier
    (« 2026.3115 ELLE Opprebais … »)."""
    name = os.path.basename(os.path.normpath(folder or ""))
    if not folder or not os.path.isdir(folder):
        raise SecuriteError("Dossier introuvable.")
    if not DOSSIER_RE.match(name):
        raise SecuriteError(
            f"« {name} » n'est pas un dossier de dossier (attendu : « 2026.xxxx ELLE/ELLS … »).\n\n"
            "Par sécurité, le programme ne travaille que dans le dossier d'un n° précis, "
            "jamais à la racine de la Dropbox ou dans un dossier de tranche.")
    return name


def backup_file(path):
    """Copie de sauvegarde (hors Dropbox) d'un fichier qu'on va écraser.
    Retourne le chemin de la copie."""
    if not os.path.exists(path):
        return None
    stamp = datetime.now().strftime("%Y-%m-%d_%Hh%M%S")
    dossier = os.path.basename(os.path.dirname(path))[:60]
    target_dir = backup_dir() / f"{stamp} {dossier}"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / os.path.basename(path)
    shutil.copy2(path, target)
    if os.path.getsize(target) != os.path.getsize(path):
        raise SecuriteError(f"La sauvegarde de « {os.path.basename(path)} » a échoué : rien n'a été modifié.")
    _purge_old()
    return str(target)


def _purge_old():
    """Ne garde que les sauvegardes des 60 derniers jours (uniquement dans
    le dossier de sauvegarde du programme, jamais ailleurs)."""
    limit = time.time() - KEEP_DAYS * 86400
    root = backup_dir()
    for d in root.iterdir():
        try:
            if d.is_dir() and d.stat().st_mtime < limit and d.parent == root:
                shutil.rmtree(d)
        except Exception:  # noqa: BLE001
            pass


def temp_path_for(path):
    base, ext = os.path.splitext(path)
    return f"{base}.~enregistrement{ext}"


def commit_temp(temp, final):
    """Remplace `final` par `temp` en une seule opération."""
    os.replace(temp, final)


def ensure_not_locked(path):
    if os.path.exists(path):
        try:
            with open(path, "r+b"):
                pass
        except OSError as exc:
            raise SecuriteError(
                f"« {os.path.basename(path)} » est ouvert (Word, Excel ou synchronisation Dropbox) : "
                "fermez-le puis recommencez. Rien n'a été modifié.") from exc
