"""Remplit les balises [[Tag]] d'un gabarit .docx en travaillant directement
sur le XML brut de word/document.xml, en reconstituant le texte à travers les
runs <w:t> successifs. Nécessaire car Word peut scinder une balise en
plusieurs runs (vérifié sur les gabarits réels de Georges : [[Localité]]
apparaît parfois éclaté à cause de la correction orthographique), donc une
simple recherche/remplacement sur un seul <w:t> ne suffit pas.

Aucune dépendance externe : uniquement la bibliothèque standard (zipfile, re).
"""
import re
import zipfile

_RUN_RE = re.compile(r"<w:t\b([^>]*?)(/>|>(.*?)</w:t>)", re.DOTALL)
_TAG_RE = re.compile(r"\[\[([^\[\]<>]{1,80}?)\]\]")


def _xml_escape(s) -> str:
    s = "" if s is None else str(s)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _find_text_runs(xml: str):
    runs = []
    for m in _RUN_RE.finditer(xml):
        attrs = m.group(1)
        self_closing = m.group(2) == "/>"
        if self_closing:
            content = ""
            content_start = content_end = m.end(1)
        else:
            content = m.group(3)
            content_start, content_end = m.start(3), m.end(3)
        runs.append(
            {
                "start": m.start(),
                "end": m.end(),
                "content_start": content_start,
                "content_end": content_end,
                "self_closing": self_closing,
                "attrs": attrs,
                "content": content,
            }
        )
    return runs


def _build_run_tag(attrs: str, content: str) -> str:
    a = attrs
    if "xml:space" not in a:
        a = a + ' xml:space="preserve"'
    return f"<w:t{a}>{content}</w:t>"


def fill_document_xml(xml: str, scalars: dict, positional: dict):
    """Retourne (nouveau_xml, balises_inconnues)."""
    runs = _find_text_runs(xml)
    if not runs:
        return xml, []

    run_ranges = []
    pos = 0
    concat_parts = []
    for run in runs:
        content = run["content"]
        g_start = pos
        concat_parts.append(content)
        pos += len(content)
        run_ranges.append({"run": run, "g_start": g_start, "g_end": pos, "content": content})
    concatenated = "".join(concat_parts)

    # Une valeur positionnelle peut être une fonction recevant le nombre
    # d'occurrences de la balise dans le gabarit (gabarits de versions
    # différentes : la même balise n'y apparaît pas le même nombre de fois).
    occurrences = {}
    for m in _TAG_RE.finditer(concatenated):
        occurrences[m.group(1)] = occurrences.get(m.group(1), 0) + 1
    positional = {k: (v(occurrences.get(k, 0)) if callable(v) else v) for k, v in positional.items()}

    counters = {}
    unknown = set()
    matches = []
    for m in _TAG_RE.finditer(concatenated):
        tag_name = m.group(1)
        ms, me = m.start(), m.end()
        if tag_name in scalars:
            value = scalars[tag_name]
        elif tag_name in positional:
            counters[tag_name] = counters.get(tag_name, 0) + 1
            idx = counters[tag_name]
            arr = positional[tag_name]
            value = arr[idx - 1] if idx <= len(arr) else ""
        else:
            unknown.add(tag_name)
            value = ""
        # Le gabarit imprime parfois déjà l'unité après la balise
        # (« [[Superficie pièces  collectives]] m² ») : on ne la répète pas
        # (GB 25/11/2026, « 17,4 m² m² »).
        if isinstance(value, str) and value.endswith(" m²") and re.match(r"\s*m²", concatenated[me:me + 6]):
            value = value[:-3]
        matches.append((ms, me, _xml_escape(value)))

    if not matches:
        return xml, []

    out_parts = []
    cursor_in_xml = 0
    for rr in run_ranges:
        run = rr["run"]
        g_start, g_end, content = rr["g_start"], rr["g_end"], rr["content"]
        relevant = [mm for mm in matches if mm[0] < g_end and mm[1] > g_start]
        if not relevant:
            new_content = content
        else:
            parts = []
            cursor = g_start
            for ms, me, value in relevant:
                seg_start = max(ms, g_start)
                seg_end = min(me, g_end)
                if seg_start > cursor:
                    parts.append(content[cursor - g_start : seg_start - g_start])
                if g_start <= ms < g_end:
                    parts.append(value)
                cursor = seg_end
            if cursor < g_end:
                parts.append(content[cursor - g_start : g_end - g_start])
            new_content = "".join(parts)

        out_parts.append(xml[cursor_in_xml : run["start"]])
        if new_content == "" and run["self_closing"]:
            out_parts.append(xml[run["start"] : run["end"]])
        else:
            out_parts.append(_build_run_tag(run["attrs"], new_content))
        cursor_in_xml = run["end"]
    out_parts.append(xml[cursor_in_xml:])
    return "".join(out_parts), sorted(unknown)


# Marqueur invisible ajouté par le moteur derrière une valeur qui porte déjà
# son unité (« 5 étudiants ») : le « personne(s) » imprimé juste après dans
# le même paragraphe du gabarit est alors retiré (GB 25/11/2026).
UNIT_MARK = "\u2063"
_UNIT_RE = re.compile(UNIT_MARK + r"((?:(?!</w:p>).)*?)\s?personne\(s\)", re.S)


def drop_unit_after_marker(xml: str) -> str:
    xml = _UNIT_RE.sub(lambda mt: mt.group(1), xml)
    return xml.replace(UNIT_MARK, "")


TRAVAUX_TAG = "Travaux_realises"
_REM_TAGS = ("Rem garde-corps", "Rem divers", "Rem Electricité", "Rem Humidité", "Rem ventilation", "Remarques")
_PARA_RE = re.compile(r"<w:p[ >].*?</w:p>", re.S)


def _para_text(p_xml):
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", p_xml))


def inject_travaux_paragraph(xml):
    """Ajoute, juste sous le dernier paragraphe de remarques du gabarit
    ([[Rem garde-corps]], à défaut une autre balise de remarques), un
    paragraphe [[Travaux_realises]] reprenant sa mise en forme. Ne fait rien
    si le gabarit contient déjà cette balise ou n'a aucune balise de
    remarques."""
    if TRAVAUX_TAG in re.sub(r"<[^>]+>", "", xml):
        return xml
    best = None
    for m in _PARA_RE.finditer(xml):
        text = _para_text(m.group(0))
        for rank, tag in enumerate(_REM_TAGS):
            if "[[" + tag + "]]" in text:
                if best is None or rank < best[0] or (rank == best[0] and m.end() > best[1].end()):
                    best = (rank, m)
                break
    if best is None:
        return xml
    para = best[1].group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", para, re.S)
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", para[ppr.end() if ppr else 0:], re.S)
    new_p = ("<w:p>" + (ppr.group(0) if ppr else "") + "<w:r>" + (rpr.group(0) if rpr else "")
             + '<w:t xml:space="preserve">[[' + TRAVAUX_TAG + "]]</w:t></w:r></w:p>")
    end = best[1].end()
    return xml[:end] + new_p + xml[end:]


def fill_docx(template_path: str, output_path: str, scalars: dict, positional: dict):
    """Remplit un gabarit .docx et écrit le résultat. Retourne la liste des
    balises [[...]] trouvées dans le gabarit mais absentes des données
    (laissées vides)."""
    with zipfile.ZipFile(template_path, "r") as zin:
        infos = zin.infolist()
        contents = {info.filename: zin.read(info.filename) for info in infos}

    doc_xml = contents["word/document.xml"].decode("utf-8")
    if str(scalars.get(TRAVAUX_TAG) or "").strip():
        doc_xml = inject_travaux_paragraph(doc_xml)
    new_xml, unknown_tags = fill_document_xml(doc_xml, scalars, positional)
    new_xml = drop_unit_after_marker(new_xml)
    contents["word/document.xml"] = new_xml.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            if info.filename.endswith("/"):
                continue
            zout.writestr(info.filename, contents[info.filename])

    return unknown_tags
