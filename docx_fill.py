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


def fill_docx(template_path: str, output_path: str, scalars: dict, positional: dict):
    """Remplit un gabarit .docx et écrit le résultat. Retourne la liste des
    balises [[...]] trouvées dans le gabarit mais absentes des données
    (laissées vides)."""
    with zipfile.ZipFile(template_path, "r") as zin:
        infos = zin.infolist()
        contents = {info.filename: zin.read(info.filename) for info in infos}

    doc_xml = contents["word/document.xml"].decode("utf-8")
    new_xml, unknown_tags = fill_document_xml(doc_xml, scalars, positional)
    contents["word/document.xml"] = new_xml.encode("utf-8")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            if info.filename.endswith("/"):
                continue
            zout.writestr(info.filename, contents[info.filename])

    return unknown_tags
