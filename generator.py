"""Orchestration : sélection du gabarit, lecture Excel, regroupement en
unités, génération des .docx (et export PDF optionnel si Word est présent).
"""
import os
import re
from pathlib import Path

import engine
import docx_fill
import xlsx_read
import plans

TEMPLATE_FILENAMES = {
    "1A": {
        "collectif": "Annexe 1A - GABARIT LOGEMENT COLLECTIF.docx",
        "individuel": "Annexe 1A - GABARIT LOGEMENT INDIVIDUEL.docx",
    },
    "1B": {
        "collectif": "Annexe 1B - GABARIT LOGEMENT COLLECTIF.docx",
        "individuel": "Annexe 1B - GABARIT LOGEMENT INDIVIDUEL.docx",
    },
    "1C": {
        "collectif": "Annexe 1C - GABARIT LOGEMENT COLLECTIF.docx",
        "individuel": "Annexe 1C - GABARIT LOGEMENT INDIVIDUEL.docx",
    },
    "1D": {
        "collectif": "Annexe 1D - GABARIT LOGEMENT COLLECTIF.docx",
        "individuel": "Annexe 1D - GABARIT LOGEMENT INDIVIDUEL.docx",
    },
}


# Attestations de conformité 2A (une par unité, à côté de chaque rapport)
ATTESTATION_FILENAMES = {
    "collectif": "Attestation 2A - GABARIT LOGEMENT COLLECTIF.docx",
    "individuel": "Attestation 2A - GABARIT LOGEMENT INDIVIDUEL.docx",
}


_NUM_RE = re.compile(r"(20\d\d)[.\-_ ](\d{4})")


def dossier_number(dossier="", *paths):
    """N° « 2026.NNNN » : celui saisi, sinon trouvé dans le chemin de l'Excel,
    en partant du dossier le plus profond. Les dossiers de tranche
    (« 2026.1000 à 2026.1099 ») sont ignorés (GB 25/11 : 2026.1039 sortait
    en 2026.1000)."""
    mm = _NUM_RE.search(dossier or "")
    if mm:
        return f"{mm.group(1)}.{mm.group(2)}"
    for txt in paths:
        parts = re.split(r"[\\/]", txt or "")
        for part in reversed(parts):
            if re.search(r"\d{4}\s+[àa]\s+20\d\d", part):
                continue  # dossier de tranche
            mm = _NUM_RE.search(part)
            if mm:
                return f"{mm.group(1)}.{mm.group(2)}"
    return (dossier or "").strip()


def _lp(path):
    """Chemin utilisable au-delà de 260 caractères sous Windows (préfixe \\\\?\\)."""
    if os.name != "nt" or not path:
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\") or len(p) < 240:
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def report_name(unit, m, dossier, initiales="GB"):
    """« 2026.NNNN PL <Commune> <Rue> <n°>.<bte> GB » (sans .docx), comme les
    attestations (GB 25/11/2026)."""
    return attestation_name(unit, m, dossier, initiales, kind="PL")


def attestation_name(unit, m, dossier, initiales="GB", kind="Attestation PL"):
    """« 2026.NNNN Attestation PL <Commune> <Rue> <n°>.<bte> GB » (sans .docx)."""
    row = unit["commun_row"] if unit["format"] == "collectif" else unit["row"]
    commune = str(engine.cell(row, m, "localite") or "").strip()
    rue = str(engine.cell(row, m, "adresse") or "").strip()
    n = engine.format_val(engine.cell(row, m, "numero"))
    bte_v = engine.cell(row, m, "bte")
    bte = str(bte_v).strip() if bte_v is not None else ""
    num = f"{n}.{bte}" if bte else n
    parts = [dossier, kind, commune, rue, num, initiales]
    name = " ".join(p for p in parts if p)
    return re.sub(r'[\\/:*?"<>|]+', "-", re.sub(r"\s+", " ", name)).strip()


def _free_path(path):
    """Ne remplace jamais une attestation existante : « … v2.docx », « … v3 »…"""
    if not os.path.exists(_lp(path)):
        return path
    base, ext = os.path.splitext(path)
    k = 2
    while os.path.exists(_lp(f"{base} v{k}{ext}")):
        k += 1
    return f"{base} v{k}{ext}"


class MissionError(Exception):
    pass


def check_gabarits_folder(folder: str):
    """Retourne la liste des 8 fichiers attendus avec leur statut trouvé/absent."""
    status = []
    for annexe, variants in TEMPLATE_FILENAMES.items():
        for fmt, fname in variants.items():
            path = os.path.join(folder or "", fname)
            status.append({"annexe": annexe, "format": fmt, "filename": fname, "found": os.path.isfile(path)})
    return status


def load_mission(excel_path: str):
    """Lit le fichier Excel et retourne (format, unites, map, unmatched, ignored_rows, extra_scalars).

    extra_scalars regroupe les valeurs trouvées dans un éventuel petit
    tableau "étiquette / valeur" d'identification du bailleur et/ou de
    l'expert, où qu'il se trouve dans le classeur (une seule fois par
    fichier, appliqué à tous les logements qu'il contient)."""
    rows = xlsx_read.read_first_sheet(excel_path)
    if not rows or len(rows) < 2:
        raise MissionError("Le fichier ne contient pas de données exploitables.")
    header = rows[0]
    m, unmatched, missing = engine.build_header_map(header)
    if missing:
        raise MissionError(
            "Colonnes obligatoires introuvables dans l'en-tête : "
            + ", ".join(missing)
            + ". Vérifiez les intitulés de colonnes du fichier Excel."
        )

    fmt = engine.detect_format(rows, m, 0)
    if not fmt:
        raise MissionError(
            "Impossible de déterminer le format des données "
            "(aucune ligne exploitable trouvée en colonne « Type Logement »)."
        )

    ignored_rows = 0
    if fmt == "collectif":
        units = engine.build_units_collectif(rows, m, 0)
        consumed = sum(1 + len(u["chambres"]) for u in units)
        total_non_empty = sum(
            1 for row in rows[1:] if not engine.is_empty(engine.cell(row, m, "typeLogement")) and str(engine.cell(row, m, "typeLogement")).strip()
        )
        ignored_rows = max(0, total_non_empty - consumed)
    else:
        units = engine.build_units_individuel(rows, m, 0)

    if not units:
        raise MissionError("Aucun logement détecté dans ce fichier.")

    all_sheets_rows = xlsx_read.read_all_sheets(excel_path)
    extra_scalars = engine.scan_bailleur_expert(all_sheets_rows)

    return fmt, units, m, unmatched, ignored_rows, extra_scalars


TRAVAUX_PHRASE = "Le bailleur a procédé aux réparations des points cités ci-dessus."


def generate_reports(
    annexe,
    units,
    m,
    gabarits_folder,
    output_folder,
    export_pdf=False,
    progress_cb=None,
    extra_scalars=None,
    plan_paths=None,
    should_stop=None,
    attestations=False,
    reports=True,
    dossier="",
    initiales="GB",
    travaux_realises=False,
):
    """Génère un .docx (et .pdf si demandé) par unité. Retourne une liste de
    dicts {file_base, docx_path, pdf_path, unknown_tags, error}.

    plan_paths : liste optionnelle de chemins (images et/ou PDF) formant un
    même jeu de plans, ajouté en pages finales de CHAQUE rapport généré ; le
    nombre total de pages est aussi renseigné dans la balise
    [[Nbre_annexes]] (via extra_scalars, sans écraser une valeur déjà
    présente pour d'autres usages).

    should_stop : fonction optionnelle, appelée avant chaque logement,
    retournant True pour arrêter proprement la génération (par exemple
    depuis un bouton "Arrêter" dans l'interface). Le logement en cours de
    traitement n'est jamais interrompu à moitié : l'arrêt prend effet entre
    deux logements, jamais pendant l'écriture de l'un d'eux. Les résultats
    déjà produits jusque-là sont retournés normalement (liste plus courte
    que le nombre total de logements)."""
    os.makedirs(output_folder, exist_ok=True)
    results = []

    rendered_pages = plans.render_pages(plan_paths) if (plan_paths and reports) else []
    scalars_with_plans = dict(extra_scalars or {})
    # Case « travaux réalisés » : phrase ajoutée sous les remarques.
    scalars_with_plans[docx_fill.TRAVAUX_TAG] = TRAVAUX_PHRASE if travaux_realises else ""
    # Le regroupement des plans dépend de la taille utile de la page du
    # gabarit : on la lit une fois sur le gabarit du lot.
    plans_area = None
    if rendered_pages and units:
        plans_area = plans.usable_area(os.path.join(gabarits_folder, TEMPLATE_FILENAMES[annexe][units[0]["format"]]))
    scalars_with_plans["Nbre_annexes"] = str(len(plans.pack_pages(rendered_pages, plans_area)))

    total = len(units)
    used_file_bases = {}
    for i, unit in enumerate(units):
        if should_stop and should_stop():
            break

        label = engine.unit_label(unit, m)
        if progress_cb:
            progress_cb(i + 1, total, label)

        fmt = unit["format"]
        template_name = TEMPLATE_FILENAMES[annexe][fmt]
        template_path = os.path.join(gabarits_folder, template_name)

        # Nom du rapport : « 2026.NNNN PL <Commune> <Rue> <n°>.<bte> GB ».
        label = dict(label)
        label["file_base"] = report_name(unit, m, dossier, initiales) or label["file_base"]
        entry = {"file_base": label["file_base"], "docx_path": None, "pdf_path": None, "unknown_tags": [], "error": None}

        # Une ligne sans adresse/n°/bte exploitable donnerait un nom de
        # fichier générique ("Rapport___") : on préfère signaler clairement
        # l'erreur plutôt que d'écrire plusieurs rapports différents sous
        # exactement le même nom (ce qui peut ensuite provoquer des erreurs
        # de permission si Word garde le fichier ouvert après une conversion
        # PDF ratée sur l'un d'eux).
        if not (label["adresse"] or label["n"] or label["bte"]):
            entry["error"] = (
                "Ligne ignorée : adresse, n° et boîte tous vides dans l'Excel "
                "(impossible de déterminer le nom du rapport)."
            )
            results.append(entry)
            continue

        # Si deux logements produisent malgré tout le même nom de fichier
        # (adresse/n°/bte identiques), on suffixe pour ne jamais écraser un
        # rapport déjà généré dans ce même lot.
        base = label["file_base"]
        if base in used_file_bases:
            used_file_bases[base] += 1
            label = dict(label)
            label["file_base"] = f"{base}_{used_file_bases[base]}"
            entry["file_base"] = label["file_base"]
        else:
            used_file_bases[base] = 1

        if not reports:
            # « Attestations seules » : aucun rapport, seulement l'attestation 2A.
            entry["attest_only"] = True
            _make_attestation(entry, unit, m, gabarits_folder, output_folder, export_pdf, scalars_with_plans,
                              dossier, initiales)
            if entry.get("attest_error"):
                entry["error"] = "Attestation : " + entry["attest_error"]
            results.append(entry)
            continue

        if not os.path.isfile(template_path):
            entry["error"] = f"Gabarit introuvable : {template_name}"
            results.append(entry)
            continue

        if fmt == "collectif":
            scalars, positional = engine.build_data_spec_collectif(unit, m, scalars_with_plans)
        else:
            scalars, positional = engine.build_data_spec_individuel(unit, m, scalars_with_plans)

        docx_path = os.path.join(output_folder, label["file_base"] + ".docx")
        try:
            unknown_tags = docx_fill.fill_docx(template_path, _lp(docx_path), scalars, positional)
            entry["docx_path"] = docx_path
            entry["unknown_tags"] = unknown_tags
        except Exception as exc:  # noqa: BLE001
            entry["error"] = str(exc)
            results.append(entry)
            continue

        if rendered_pages:
            try:
                plans.append_pages_to_docx(_lp(docx_path), rendered_pages, plans_area)
            except Exception as exc:  # noqa: BLE001
                entry["error"] = f"Plans non ajoutés : {exc}"
                results.append(entry)
                continue

        if export_pdf:
            try:
                pdf_path = _export_pdf(docx_path)
                entry["pdf_path"] = pdf_path
            except Exception as exc:  # noqa: BLE001
                entry["error"] = f"PDF non généré : {exc}"

        if attestations:
            _make_attestation(entry, unit, m, gabarits_folder, output_folder, export_pdf, scalars_with_plans,
                              dossier, initiales)

        results.append(entry)

    return results


def _make_attestation(entry, unit, m, gabarits_folder, output_folder, export_pdf, extra_scalars, dossier, initiales):
    """Attestation 2A de l'unité ; une erreur ici n'annule jamais le rapport
    (elle est notée dans entry["attest_error"])."""
    entry.update({"attest_docx": None, "attest_pdf": None, "attest_error": None, "attest_unknown": []})
    tpl = os.path.join(gabarits_folder, ATTESTATION_FILENAMES[unit["format"]])
    if not os.path.isfile(tpl):
        entry["attest_error"] = f"Gabarit introuvable : {ATTESTATION_FILENAMES[unit['format']]}"
        return
    try:
        scalars, positional = engine.build_data_spec_attestation(unit, m, extra_scalars)
        path = _free_path(os.path.join(output_folder, attestation_name(unit, m, dossier, initiales) + ".docx"))
        entry["attest_unknown"] = docx_fill.fill_docx(tpl, _lp(path), scalars, positional)
        entry["attest_docx"] = path
    except Exception as exc:  # noqa: BLE001
        entry["attest_error"] = str(exc)
        return
    if export_pdf:
        try:
            entry["attest_pdf"] = _export_pdf(path)
        except Exception as exc:  # noqa: BLE001
            entry["attest_error"] = f"PDF de l'attestation non généré : {exc}"


def _export_pdf(docx_path: str) -> str:
    """Convertit un .docx en PDF via Word (nécessite Word installé sur
    Windows). Lève une exception si docx2pdf n'est pas disponible ou échoue —
    ce n'est jamais bloquant pour la génération du .docx lui-même."""
    from docx2pdf import convert  # import différé : dépendance optionnelle

    pdf_path = str(Path(docx_path).with_suffix(".pdf"))
    if len(os.path.abspath(pdf_path)) < 240:
        convert(docx_path, pdf_path)
        return pdf_path
    # Chemin trop long pour Word : conversion dans un dossier temporaire
    # court, puis déplacement du PDF à côté du .docx.
    import shutil
    import tempfile

    with tempfile.TemporaryDirectory(prefix="pl_") as tmp:
        tmp_docx = os.path.join(tmp, "doc.docx")
        tmp_pdf = os.path.join(tmp, "doc.pdf")
        shutil.copyfile(_lp(docx_path), tmp_docx)
        convert(tmp_docx, tmp_pdf)
        shutil.move(tmp_pdf, _lp(pdf_path))
    return pdf_path
