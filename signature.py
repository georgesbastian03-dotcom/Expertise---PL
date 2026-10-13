"""Emplacement de l'image de signature de l'expert.

La signature n'est jamais publiée avec le programme (dépôt GitHub public) :
elle est gardée sur ce PC, dans %APPDATA%\\RapportsPermisLocation\\, choisie
une seule fois par Georges (menu Aide → « Ma signature… », ou à la première
génération qui en a besoin). Une copie locale « assets/signature.jpg » à côté
du programme est aussi acceptée.
"""
import os
import shutil

import config

APP_DIR = os.path.dirname(os.path.abspath(__file__))
EXTS = (".png", ".jpg", ".jpeg")


def _stored_candidates():
    d = config._config_dir()
    return [os.path.join(d, "signature" + e) for e in EXTS]


def signature_path():
    for p in _stored_candidates():
        if os.path.exists(p):
            return p
    for e in EXTS:
        p = os.path.join(APP_DIR, "assets", "signature" + e)
        if os.path.exists(p):
            return p
    return None


def set_signature(source_path):
    ext = os.path.splitext(source_path)[1].lower()
    if ext not in EXTS:
        raise ValueError("Choisissez une image PNG ou JPG.")
    for p in _stored_candidates():
        if os.path.exists(p):
            os.remove(p)
    target = os.path.join(config._config_dir(), "signature" + (".jpg" if ext == ".jpeg" else ext))
    shutil.copy2(source_path, target)
    return target
