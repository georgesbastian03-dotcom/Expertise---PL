"""Vérification et installation des mises à jour depuis les Releases GitHub.

Fonctionne uniquement avec la bibliothèque standard (urllib, zipfile) : aucune
dépendance supplémentaire. Toute erreur réseau est avalée silencieusement par
check_for_update — l'absence de connexion ne doit jamais empêcher l'usage
normal du programme.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile

# Dépôt GitHub où sont publiées les nouvelles versions (Releases, avec un
# fichier .zip en pièce jointe à chaque release).
REPO = "georgesbastian03-dotcom/Expertise---PL"
API_URL = f"https://api.github.com/repos/{REPO}/releases/latest"

# Version de CE code. À incrémenter à chaque nouvelle version publiée sur
# GitHub (le tag de la release doit être la même valeur, avec ou sans "v").
APP_VERSION = "1.1.0"


def _parse_version(v):
    v = (v or "").strip()
    if v[:1] in ("v", "V"):
        v = v[1:]
    parts = []
    for p in v.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def check_for_update(timeout=6):
    """Retourne {"tag", "url", "notes"} si une version plus récente est
    disponible sur GitHub, sinon None. Ne lève jamais d'exception."""
    try:
        req = urllib.request.Request(
            API_URL,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "RapportsPermisLocation",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        tag = data.get("tag_name", "")
        if not tag or _parse_version(tag) <= _parse_version(APP_VERSION):
            return None

        zip_url = None
        for asset in data.get("assets", []):
            name = (asset.get("name") or "").lower()
            if name.endswith(".zip"):
                zip_url = asset.get("browser_download_url")
                break
        if not zip_url:
            return None

        return {"tag": tag, "url": zip_url, "notes": (data.get("body") or "").strip()}
    except Exception:
        return None


def app_dir():
    """Dossier contenant les fichiers de l'application (celui de ce module,
    toujours à côté de main.py)."""
    return os.path.dirname(os.path.abspath(__file__))


def _extract_stripping_common_root(zip_path, dest_dir):
    with zipfile.ZipFile(zip_path) as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        common_prefix = None
        if names:
            first = names[0]
            if "/" in first:
                candidate = first.split("/", 1)[0] + "/"
                if all(n.startswith(candidate) for n in names):
                    common_prefix = candidate
        zf.extractall(dest_dir)
    return os.path.join(dest_dir, common_prefix) if common_prefix else dest_dir


def download_and_install(zip_url, progress_cb=None):
    """Télécharge et installe la mise à jour par-dessus le dossier actuel de
    l'application. N'affecte jamais votre dossier de gabarits ni vos réglages
    (stockés séparément, hors de ce dossier)."""
    target_dir = app_dir()

    def notify(msg):
        if progress_cb:
            progress_cb(msg)

    with tempfile.TemporaryDirectory(prefix="maj_rapports_") as tmp:
        zip_path = os.path.join(tmp, "update.zip")
        notify("Téléchargement de la mise à jour…")
        req = urllib.request.Request(zip_url, headers={"User-Agent": "RapportsPermisLocation"})
        with urllib.request.urlopen(req, timeout=120) as resp, open(zip_path, "wb") as f:
            shutil.copyfileobj(resp, f)

        notify("Installation de la mise à jour…")
        extract_dir = os.path.join(tmp, "extracted")
        src_root = _extract_stripping_common_root(zip_path, extract_dir)

        for root, _dirs, files in os.walk(src_root):
            rel = os.path.relpath(root, src_root)
            dest_root = target_dir if rel == "." else os.path.join(target_dir, rel)
            os.makedirs(dest_root, exist_ok=True)
            for name in files:
                shutil.copy2(os.path.join(root, name), os.path.join(dest_root, name))

    notify("Mise à jour installée.")


def relaunch_and_exit():
    """Démarre une nouvelle instance de l'application puis termine celle-ci."""
    python = sys.executable
    main_py = os.path.join(app_dir(), "main.py")
    creationflags = 0
    if sys.platform.startswith("win"):
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([python, main_py], cwd=app_dir(), creationflags=creationflags)
    os._exit(0)
