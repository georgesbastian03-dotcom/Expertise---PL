"""Montants en toutes lettres, en français de Belgique (septante, nonante),
au format utilisé dans les PV de sortie locative :

    1377.18 -> "MILLE TROIS CENT SEPTANTE SEPT EUROS DIX HUIT CENTIMES"

Les mots sont séparés par des espaces (pas de traits d'union), en majuscules,
comme dans le PV du dossier 2026.2246.
"""

_UNITES = [
    "zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf",
    "dix", "onze", "douze", "treize", "quatorze", "quinze", "seize",
    "dix sept", "dix huit", "dix neuf",
]
_DIZAINES = {
    2: "vingt", 3: "trente", 4: "quarante", 5: "cinquante",
    6: "soixante", 7: "septante", 8: "quatre-vingt", 9: "nonante",
}


def _moins_de_cent(n):
    if n < 20:
        return _UNITES[n]
    d, u = divmod(n, 10)
    base = _DIZAINES[d]
    if u == 0:
        return "quatre-vingts" if d == 8 else base
    if u == 1 and d != 8:
        return f"{base} et un"
    return f"{base} {_UNITES[u]}"


def _moins_de_mille(n):
    c, reste = divmod(n, 100)
    parts = []
    if c:
        if c == 1:
            parts.append("cent")
        else:
            parts.append(f"{_UNITES[c]} cent" + ("s" if reste == 0 else ""))
    if reste:
        parts.append(_moins_de_cent(reste))
    return " ".join(parts)


def entier_en_lettres(n):
    n = int(n)
    if n == 0:
        return "zéro"
    parts = []
    millions, reste = divmod(n, 1_000_000)
    milliers, unites = divmod(reste, 1000)
    if millions:
        parts.append("un million" if millions == 1 else f"{_moins_de_mille(millions)} millions")
    if milliers:
        if milliers == 1:
            parts.append("mille")
        else:
            txt = _moins_de_mille(milliers)
            # « deux cents » devient « deux cent » devant « mille »
            if txt.endswith("cents"):
                txt = txt[:-1]
            if txt.endswith("quatre-vingts"):
                txt = txt[:-1]
            parts.append(f"{txt} mille")
    if unites:
        parts.append(_moins_de_mille(unites))
    return " ".join(parts)


def montant_en_lettres(montant):
    """Retourne le montant en lettres majuscules, ex. 'MILLE TROIS CENT
    SEPTANTE SEPT EUROS DIX HUIT CENTIMES'."""
    cents_total = int(round(float(montant) * 100))
    euros, centimes = divmod(abs(cents_total), 100)
    txt = entier_en_lettres(euros) + (" euro" if euros <= 1 else " euros")
    if centimes:
        txt += " " + entier_en_lettres(centimes) + (" centime" if centimes == 1 else " centimes")
    txt = txt.replace("-", " ")
    return txt.upper()


def montant_en_chiffres(montant):
    """1377.18 -> '1.377,18 €' (format belge)."""
    s = f"{float(montant):,.2f}"  # 1,377.18
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


if __name__ == "__main__":
    for v in (1377.18, 0.5, 1, 71, 80, 81, 91, 200, 1200, 80000, 2021.01, 99.99):
        print(v, montant_en_chiffres(v), "|", montant_en_lettres(v))
