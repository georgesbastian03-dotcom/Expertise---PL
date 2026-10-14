"""Paramètres persistants de l'application (dossier des gabarits, dernière
annexe utilisée, derniers dossiers parcourus...). Enregistrés dans un fichier
JSON sous %APPDATA%\\RapportsPermisLocation (ou ~/.rapports_permis_location
si APPDATA n'existe pas, ex. hors Windows) — indépendant de l'emplacement du
programme, comme le faisait SaveSetting/GetSetting dans la macro VBA.
"""
import json
import os
from pathlib import Path

DEFAULTS = {
    "gabarits_folder": "",
    "last_annexe": "1A",
    "last_excel_dir": "",
    "last_output_dir": "",
    "export_pdf": False,
    # --- Onglets "État des lieux" (entrée / sortie), 100 % locaux ---------
    "last_entree_parent": "",
    "last_sortie_parent": "",
    "last_entree_batch": "",
    "mes_initiales": "GB",
}


def _config_dir() -> Path:
    appdata = os.environ.get("APPDATA")
    base = Path(appdata) if appdata else Path.home()
    d = base / "RapportsPermisLocation"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _config_path() -> Path:
    return _config_dir() / "config.json"


def load_config() -> dict:
    p = _config_path()
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            merged = dict(DEFAULTS)
            merged.update(data)
            # Ancienne version (lecture IA) : la clé API n'est plus utilisée.
            merged.pop("anthropic_api_key", None)
            return merged
        except Exception:
            pass
    return dict(DEFAULTS)


def save_config(cfg: dict) -> None:
    p = _config_path()
    try:
        p.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass  # les préférences ne sont jamais bloquantes
