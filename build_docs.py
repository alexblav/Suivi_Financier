#!/usr/bin/env python3
"""Convertit README.md en README.html (fichier autonome, sans dépendance en ligne).

Usage :  python build_docs.py            (pip install markdown, une seule fois)

- README.md est la SOURCE ; README.html en est la conversion : ne pas l'éditer à la main.
- Les blocs ```mermaid (sous-ensemble : flowchart TB/LR, noeuds [..] [(..)] {..},
  fleches -->, -->|"texte"|, -.-) sont dessinés en SVG par le mini-moteur ci-dessous,
  pour que le HTML s'ouvre hors connexion.
- Le sommaire latéral est généré à partir des titres de niveau 2 et 3.
"""
import html
import re
import sys
from pathlib import Path

import markdown

RACINE = Path(__file__).resolve().parent
SOURCE = RACINE / "README.md"
CIBLE = RACINE / "README.html"

# ---------------------------------------------------------------------------
# Mermaid (sous-ensemble) -> SVG
# ---------------------------------------------------------------------------
CH_W, LIGNE_H = 7.4, 17
GAP_COUCHE, GAP_NOEUD = 70, 36


class Noeud:
    def __init__(self, ident):
        self.id = ident
        self.texte = ident
        self.forme = "rect"
        self.couche = 0
        self.ordre = 0.0
        self.x = self.y = self.w = self.h = 0.0
        self.lignes = []


def decouper(texte, largeur=26):
    lignes = []
    for brut in re.split(r"<br\s*/?>", texte):
        mots, cur = brut.split(), ""
        for m in mots:
            if cur and len(cur) + 1 + len(m) > largeur:
                lignes.append(cur)
                cur = m
            else:
                cur = (cur + " " + m).strip()
        lignes.append(cur)
    return lignes or [""]


NOEUD_RE = re.compile(
    r'\s*([A-Za-z_]\w*)(?:\[\("(.*?)"\)\]|\["(.*?)"\]|\{"(.*?)"\})?')
FLECHE_RE = re.compile(r'\s*(-->|-\.-)(?:\|"(.*?)"\|)?')


def lire_mermaid(src):
    lignes = [l.strip() for l in src.strip().splitlines() if l.strip()]
    m = re.match(r"flowchart\s+(TB|TD|LR)", lignes[0])
    sens = "LR" if m and m.group(1) == "LR" else "TB"
    noeuds, aretes = {}, []

    def noeud(mt):
        ident = mt.group(1)
        n = noeuds.setdefault(ident, Noeud(ident))
        if mt.group(2) is not None:
            n.texte, n.forme = mt.group(2), "db"
        elif mt.group(3) is not None:
            n.texte = mt.group(3)
        elif mt.group(4) is not None:
            n.texte, n.forme = mt.group(4), "diamond"
        return n

    for l in lignes[1:]:
        pos = 0
        mt = NOEUD_RE.match(l, pos)
        if not mt:
            continue
        courant, pos = noeud(mt), mt.end()
        while True:
            fl = FLECHE_RE.match(l, pos)
            if not fl:
                break
            suivant = NOEUD_RE.match(l, fl.end())
            if not suivant:
                break
            cible = noeud(suivant)
            aretes.append((courant, cible, fl.group(2) or "", fl.group(1) == "-.-"))
            courant, pos = cible, suivant.end()
    return sens, noeuds, aretes


def mise_en_page(sens, noeuds, aretes):
    # dimensions
    for n in noeuds.values():
        n.lignes = decouper(n.texte, 21 if sens == "LR" else 26)
        lw = max(len(x) for x in n.lignes) * CH_W + 30
        lh = len(n.lignes) * LIGNE_H + 20
        if n.forme == "diamond":
            lw, lh = lw * 1.35 + 20, lh * 1.5 + 6
        n.w, n.h = max(lw, 110), lh

    # detection des arcs retour (DFS)
    succ = {i: [] for i in noeuds}
    for a, b, _, _ in aretes:
        succ[a.id].append(b.id)
    etat, retour = {}, set()

    def dfs(u):
        etat[u] = 1
        for v in reversed(succ[u]):
            if etat.get(v) == 1:
                retour.add((u, v))
            elif v not in etat:
                dfs(v)
        etat[u] = 2

    for i in noeuds:
        if i not in etat:
            dfs(i)
    avant = [(a, b) for a, b, _, _ in aretes if (a.id, b.id) not in retour]

    # couches : chemin le plus long, puis les sources sont rapprochees de leurs cibles
    preds = {i: [] for i in noeuds}
    succs = {i: [] for i in noeuds}
    for a, b in avant:
        preds[b.id].append(a.id)
        succs[a.id].append(b.id)
    couche = {}

    def niv(i):
        if i not in couche:
            couche[i] = 0 if not preds[i] else 1 + max(niv(p) for p in preds[i])
        return couche[i]

    for i in noeuds:
        niv(i)
    for i in noeuds:
        if not preds[i] and succs[i]:
            couche[i] = min(couche[s] for s in succs[i]) - 1
    for i, n in noeuds.items():
        n.couche = couche[i]

    # ordre dans chaque couche (barycentre)
    nb = max(n.couche for n in noeuds.values()) + 1
    cols = [[n for n in noeuds.values() if n.couche == c] for c in range(nb)]
    for c in cols:
        for k, n in enumerate(c):
            n.ordre = k
    for _ in range(4):
        for c in range(1, nb):
            for n in cols[c]:
                ps = [noeuds[p].ordre for p in preds[n.id]]
                if ps:
                    n.ordre = sum(ps) / len(ps)
            cols[c].sort(key=lambda n: n.ordre)
            for k, n in enumerate(cols[c]):
                n.ordre = k

    retour_tb = bool(retour)
    # coordonnees
    if sens == "TB":
        largeurs = [sum(n.w for n in c) + GAP_NOEUD * (len(c) - 1) for c in cols]
        total = max(largeurs)
        y = 20
        for c, col in enumerate(cols):
            x = 20 + (total - largeurs[c]) / 2
            hmax = max(n.h for n in col)
            for n in col:
                n.x, n.y = x, y + (hmax - n.h) / 2
                x += n.w + GAP_NOEUD
            y += hmax + GAP_COUCHE
        W, H = total + 40 + (240 if retour_tb else 60), y - GAP_COUCHE + 40
    else:
        hauteurs = [sum(n.h for n in c) + GAP_NOEUD * (len(c) - 1) for c in cols]
        total = max(hauteurs)
        x = 20
        for c, col in enumerate(cols):
            y = 20 + (total - hauteurs[c]) / 2
            wmax = max(n.w for n in col)
            for n in col:
                n.x, n.y = x + (wmax - n.w) / 2, y
                y += n.h + GAP_NOEUD
            x += wmax + 56
        W, H = x - 56 + 40, total + 40 + 60
    return W, H, retour


def point_bord(n, cote):
    cx, cy = n.x + n.w / 2, n.y + n.h / 2
    return {"b": (cx, n.y + n.h), "h": (cx, n.y), "g": (n.x, cy), "d": (n.x + n.w, cy)}[cote]


def svg_mermaid(src, idx):
    sens, noeuds, aretes = lire_mermaid(src)
    W, H, retour = mise_en_page(sens, noeuds, aretes)
    out = [f'<svg class="schema" width="{W:.0f}" viewBox="0 0 {W:.0f} {H:.0f}" role="img" '
           f'aria-label="Schéma {idx}" xmlns="http://www.w3.org/2000/svg">',
           '<defs><marker id="fl%d" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" '
           'markerHeight="8" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" '
           'class="pointe"/></marker></defs>' % idx]
    labels = []
    for a, b, lab, pointille in aretes:
        est_retour = (a.id, b.id) in retour
        if sens == "TB":
            if est_retour:
                (x1, y1), (x2, y2) = point_bord(a, "d"), point_bord(b, "d")
                dx = max(a.x + a.w, b.x + b.w) + 60
                d = f"M{x1:.0f},{y1:.0f} C{dx:.0f},{y1:.0f} {dx:.0f},{y2:.0f} {x2:.0f},{y2:.0f}"
                lx, ly = dx + len(lab) * 3.3 + 22, (y1 + y2) / 2
            else:
                (x1, y1), (x2, y2) = point_bord(a, "b"), point_bord(b, "h")
                my = (y1 + y2) / 2
                d = f"M{x1:.0f},{y1:.0f} C{x1:.0f},{my:.0f} {x2:.0f},{my:.0f} {x2:.0f},{y2:.0f}"
                lx, ly = (x1 + x2) / 2, (y1 + 3 * y2) / 4 if abs(x1 - x2) < 2 else my
        else:
            if est_retour:
                (x1, y1), (x2, y2) = point_bord(a, "b"), point_bord(b, "b")
                dy = max(a.y + a.h, b.y + b.h) + 45
                d = f"M{x1:.0f},{y1:.0f} C{x1:.0f},{dy:.0f} {x2:.0f},{dy:.0f} {x2:.0f},{y2:.0f}"
                lx, ly = (x1 + x2) / 2, dy - 6
            else:
                (x1, y1), (x2, y2) = point_bord(a, "d"), point_bord(b, "g")
                mx = (x1 + x2) / 2
                d = f"M{x1:.0f},{y1:.0f} C{mx:.0f},{y1:.0f} {mx:.0f},{y2:.0f} {x2:.0f},{y2:.0f}"
                lx, ly = mx, (y1 + y2) / 2
        cls = "arete pointille" if pointille else "arete"
        marq = "" if pointille else f' marker-end="url(#fl{idx})"'
        out.append(f'<path class="{cls}" d="{d}"{marq}/>')
        if lab:
            labels.append((lx, ly, lab))
    for n in noeuds.values():
        cx, cy = n.x + n.w / 2, n.y + n.h / 2
        if n.forme == "diamond":
            pts = f"{cx:.0f},{n.y:.0f} {n.x + n.w:.0f},{cy:.0f} {cx:.0f},{n.y + n.h:.0f} {n.x:.0f},{cy:.0f}"
            out.append(f'<polygon class="noeud decision" points="{pts}"/>')
        else:
            cls = "noeud base" if n.forme == "db" else "noeud"
            rx = 16 if n.forme == "db" else 8
            out.append(f'<rect class="{cls}" x="{n.x:.0f}" y="{n.y:.0f}" width="{n.w:.0f}" '
                       f'height="{n.h:.0f}" rx="{rx}"/>')
        top = cy - (len(n.lignes) - 1) * LIGNE_H / 2
        for k, ligne in enumerate(n.lignes):
            out.append(f'<text class="txt{" t1" if k == 0 and len(n.lignes) > 1 else ""}" '
                       f'x="{cx:.0f}" y="{top + k * LIGNE_H + 4.5:.1f}" text-anchor="middle">'
                       f'{html.escape(ligne)}</text>')
    for lx, ly, lab in labels:
        w = len(lab) * 6.6 + 12
        out.append(f'<rect class="fond-etiquette" x="{lx - w / 2:.0f}" y="{ly - 10:.0f}" '
                   f'width="{w:.0f}" height="18" rx="4"/>'
                   f'<text class="etiquette" x="{lx:.0f}" y="{ly + 3:.0f}" text-anchor="middle">'
                   f'{html.escape(lab)}</text>')
    out.append("</svg>")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Markdown -> HTML
# ---------------------------------------------------------------------------
def slug(valeur, sep="-"):
    """Même algorithme que GitHub : les liens du sommaire du .md fonctionnent tels quels."""
    v = valeur.strip().lower()
    v = re.sub(r"[^\w\- ]", "", v)
    return v.replace(" ", sep)


def formule(latex):
    t = re.sub(r"\\text\{(.*?)\}", r"\1", latex)
    t = t.replace("-", " − ").replace("=", " = ")
    return html.escape(re.sub(r"\s+", " ", t).strip())


def convertir(md_texte):
    blocs = []

    def garde(m):
        langue, corps = m.group(1), m.group(2)
        if langue == "mermaid":
            blocs.append('<figure class="diagramme">' + svg_mermaid(corps, len(blocs) + 1) + "</figure>")
        elif langue == "latex":
            blocs.append(f'<div class="formule">{formule(corps)}</div>')
        else:
            return m.group(0)
        return f"\n\n@@BLOC{len(blocs) - 1}@@\n\n"

    texte = re.sub(r"```(\w+)\n(.*?)```", garde, md_texte, flags=re.S)
    md = markdown.Markdown(extensions=["tables", "fenced_code", "toc", "sane_lists"],
                           extension_configs={"toc": {"slugify": slug}})
    corps = md.convert(texte)
    for i, b in enumerate(blocs):
        corps = corps.replace(f"<p>@@BLOC{i}@@</p>", b)
    corps = re.sub(r"<table>", '<div class="tableau"><table>', corps)
    corps = corps.replace("</table>", "</table></div>")
    return corps, md.toc_tokens


def nav(tokens):
    def rendu(liste, niveau):
        items = []
        for t in liste:
            if t["level"] == 1:
                items.append(rendu(t["children"], niveau))
                continue
            if t["level"] > 3:
                continue
            sous = rendu(t["children"], niveau + 1) if t["children"] and t["level"] < 3 else ""
            items.append(f'<li><a href="#{t["id"]}">{html.escape(t["name"])}</a>{sous}</li>')
        return "<ul>" + "".join(items) + "</ul>" if items else ""
    return rendu(tokens, 2)


CSS = """
:root{--fond:#fbfbfa;--texte:#1d2430;--doux:#5b6678;--trait:#d9dee6;--accent:#1f497d;
--accent2:#2f6db5;--carte:#ffffff;--code:#f1f3f6;--noeud:#eaf1fb;--base:#fdf3dc;--decision:#fbe9e7}
@media (prefers-color-scheme:dark){:root{--fond:#14181f;--texte:#e3e8ef;--doux:#9aa6b8;
--trait:#2d3644;--accent:#8db4e2;--accent2:#6aa3e6;--carte:#1b212b;--code:#222a36;
--noeud:#1f2d44;--base:#3a3220;--decision:#3d2623}}
*{box-sizing:border-box}html{scroll-behavior:smooth;scroll-padding-top:16px}
body{margin:0;background:var(--fond);color:var(--texte);
font:16px/1.65 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.page{display:flex;align-items:flex-start;max-width:1240px;margin:0 auto}
nav.sommaire{position:sticky;top:0;flex:0 0 290px;max-height:100vh;overflow:auto;
padding:24px 16px 24px 24px;font-size:14px;border-right:1px solid var(--trait)}
nav.sommaire .titre{font-weight:700;color:var(--accent);margin:0 0 8px;font-size:15px}
nav.sommaire ul{list-style:none;margin:0;padding:0}nav.sommaire ul ul{padding-left:14px;
border-left:1px solid var(--trait);margin:2px 0 6px 4px}
nav.sommaire a{display:block;padding:3px 6px;color:var(--doux);text-decoration:none;border-radius:6px}
nav.sommaire a:hover{background:var(--code);color:var(--accent)}
main{flex:1;min-width:0;padding:28px 36px 80px;max-width:920px}
h1{font-size:2rem;color:var(--accent);margin:.2em 0 .1em;line-height:1.2}
h2{font-size:1.5rem;color:var(--accent);margin:2.2em 0 .6em;padding-bottom:.25em;
border-bottom:2px solid var(--trait)}
h3{font-size:1.18rem;margin:1.6em 0 .4em;color:var(--accent2)}
a{color:var(--accent2)}p,li{max-width:76ch}
code{background:var(--code);padding:.1em .38em;border-radius:5px;font:.88em ui-monospace,Consolas,monospace;
word-break:break-word}
pre{background:var(--code);padding:12px 14px;border-radius:8px;overflow:auto}pre code{padding:0;background:none}
.tableau{overflow-x:auto;margin:1em 0;border:1px solid var(--trait);border-radius:8px;background:var(--carte)}
table{border-collapse:collapse;width:100%;font-size:.93rem}
th,td{padding:7px 11px;text-align:left;vertical-align:top;border-bottom:1px solid var(--trait)}
th{background:var(--code);color:var(--accent);font-weight:650;position:sticky;top:0}
tr:last-child td{border-bottom:0}
blockquote{margin:1em 0;padding:.2em 1em;border-left:4px solid var(--accent2);background:var(--code)}
.formule{margin:1em 0;padding:14px 18px;background:var(--carte);border:1px solid var(--trait);
border-left:4px solid var(--accent2);border-radius:8px;font:1.08rem ui-monospace,Consolas,monospace}
figure.diagramme{margin:1.4em 0;padding:14px;background:var(--carte);border:1px solid var(--trait);
border-radius:10px;overflow-x:auto;text-align:center}
svg.schema{max-width:100%;height:auto;min-width:340px}
svg .noeud{fill:var(--noeud);stroke:var(--accent2);stroke-width:1.4}
svg .noeud.base{fill:var(--base);stroke:#c28b1e;stroke-width:2}
svg .noeud.decision{fill:var(--decision);stroke:#c4503f}
svg .txt{fill:var(--texte);font:13px system-ui,sans-serif}svg .txt.t1{font-weight:650}
svg .arete{fill:none;stroke:var(--doux);stroke-width:1.5}svg .pointille{stroke-dasharray:5 4}
svg .pointe{fill:var(--doux)}
svg .fond-etiquette{fill:var(--carte);stroke:var(--trait)}
svg .etiquette{fill:var(--doux);font:11.5px system-ui,sans-serif}
.haut{position:fixed;right:18px;bottom:18px;background:var(--accent);color:var(--fond);
padding:8px 12px;border-radius:999px;text-decoration:none;font-size:13px;box-shadow:0 2px 8px #0003}
.meta{color:var(--doux);margin-top:0}
@media (max-width:900px){.page{display:block}nav.sommaire{position:static;max-height:none;
border-right:0;border-bottom:1px solid var(--trait);padding:16px}main{padding:20px 16px 60px}}
@media print{nav.sommaire,.haut{display:none}.page{display:block}main{max-width:none;padding:0}
h2{break-after:avoid}figure,.tableau,pre{break-inside:avoid}body{background:#fff;color:#000}}
"""


def main():
    md_texte = SOURCE.read_text(encoding="utf-8")
    corps, tokens = convertir(md_texte)
    # le premier <h1> du .md sert de titre ; le sommaire du .md reste dans le corps
    m = re.search(r"<h1[^>]*>(.*?)</h1>", corps, flags=re.S)
    titre = re.sub(r"<[^>]+>", "", m.group(1)) if m else "Suivi Financier"
    page = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(titre)}</title>
<style>{CSS}</style>
</head>
<body>
<div class="page">
<nav class="sommaire" aria-label="Sommaire"><p class="titre">Sommaire</p>{nav(tokens)}</nav>
<main>
{corps}
</main>
</div>
<a class="haut" href="#">&uarr; Haut</a>
<!-- Généré par build_docs.py à partir de README.md : ne pas modifier ce fichier à la main. -->
</body>
</html>
"""
    CIBLE.write_text(page, encoding="utf-8", newline="\n")
    print(f"{CIBLE.name} genere ({len(page) // 1024} Ko, {page.count('<svg class')} schemas).")


if __name__ == "__main__":
    sys.exit(main())
