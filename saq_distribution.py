"""
saq_distribution.py
-------------------
Récupère la distribution SAQ (nombre de succursales) pour une liste de produits.
Roule dans le Codespace GitHub : python3 saq_distribution.py

Dépendances : pip install requests openpyxl
"""

import requests
import re
import time
import json
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIG — ta liste de produits (code_saq, nom)
# ---------------------------------------------------------------------------
PRODUITS = [
    ("15544570", "Beach Day Every Day Vodka Frozen Berries"),
    ("14773",    "Mate Libre Eau Pétillante"),
    # Ajoute tes autres produits ici...
]

CACHE_FILE   = "saq_id_cache.json"   # mapping code → ID interne (réutilisé chaque semaine)
DELAI        = 1.2                    # secondes entre requêtes
# ---------------------------------------------------------------------------

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Language": "fr-CA,fr;q=0.9",
})


def load_cache() -> dict:
    if Path(CACHE_FILE).exists():
        with open(CACHE_FILE, "r") as f:
            return json.load(f)
    return {}


def save_cache(cache: dict):
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)


def get_internal_id(code_saq: str, cache: dict) -> str | None:
    """Trouve l'ID interne SAQ à partir du code SAQ via la redirection de page."""
    if code_saq in cache:
        return cache[code_saq]

    try:
        r = SESSION.get(
            f"https://www.saq.com/en/{code_saq}",
            timeout=15,
            allow_redirects=True
        )
        # L'URL finale ressemble à :
        # https://www.saq.com/en/cooler/.../15544570?srsltid=...
        # On cherche l'ID dans le HTML — présent dans les liens de store locator
        match = re.search(r'/store/locator/ajaxlist/context/product/id/(\d+)', r.text)
        if not match:
            # Fallback : chercher dans window.__initialData ou data-product-id
            match = re.search(r'"product_id"\s*:\s*"?(\d+)"?', r.text)
        if not match:
            match = re.search(r'data-product-id="(\d+)"', r.text)

        if match:
            internal_id = match.group(1)
            cache[code_saq] = internal_id
            save_cache(cache)
            return internal_id

    except requests.RequestException as e:
        print(f"  Erreur page produit : {e}")

    return None


def get_distribution(internal_id: str) -> dict:
    """Appelle l'API de distribution SAQ et retourne total + liste succursales."""
    url = (
        f"https://www.saq.com/en/store/locator/ajaxlist/context/product/id/{internal_id}"
        f"?loaded=0&latitude=45.5088&longitude=-73.5540"
    )
    try:
        r = SESSION.get(url, timeout=15)
        r.raise_for_status()
        data = r.json()
        stores = data.get("list", [])
        total  = data.get("total", len(stores))
        return {
            "total": total,
            "stores": [
                {
                    "nom":   s.get("name", ""),
                    "ville": s.get("city", ""),
                    "qty":   s.get("qty", 0),
                    "type":  s.get("additional_attributes", {}).get("type", {}).get("label", ""),
                }
                for s in stores
            ],
            "erreur": "",
        }
    except Exception as e:
        return {"total": 0, "stores": [], "erreur": str(e)}


def build_excel(resultats: list, filepath: str):
    wb = openpyxl.Workbook()

    NAVY  = "002060"
    GREY  = "F2F2F2"
    WHITE = "FFFFFF"
    GREEN = "00703C"
    RED   = "C00000"

    # ── Onglet 1 : Résumé ──────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Résumé distribution"

    ws.merge_cells("A1:E1")
    ws["A1"] = f"Distribution SAQ — {datetime.today().strftime('%d %B %Y')}"
    ws["A1"].font      = Font(name="Arial", size=13, bold=True, color=WHITE)
    ws["A1"].fill      = PatternFill("solid", start_color=NAVY)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for col, h in enumerate(["Code SAQ", "Produit", "# Succursales", "Statut", "Erreur"], 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font      = Font(name="Arial", size=10, bold=True, color=WHITE)
        c.fill      = PatternFill("solid", start_color=NAVY)
        c.alignment = Alignment(horizontal="center")

    for i, r in enumerate(resultats, 3):
        fill = PatternFill("solid", start_color=GREY if i % 2 == 0 else WHITE)
        vals = [r["code"], r["nom"], r["total"], r["statut"], r["erreur"]]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font      = Font(name="Arial", size=10)
            c.fill      = fill
            c.alignment = Alignment(horizontal="center" if col in (1, 3) else "left")

        nb = ws.cell(row=i, column=3)
        if r["total"] == 0:
            nb.font = Font(name="Arial", size=10, bold=True, color=RED)
        elif r["total"] >= 100:
            nb.font = Font(name="Arial", size=10, bold=True, color=GREEN)

    for col, w in zip("ABCDE", [14, 45, 15, 12, 40]):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ── Onglet 2 : Détail ──────────────────────────────────────────────────
    ws2 = wb.create_sheet("Détail succursales")
    for col, h in enumerate(["Code SAQ", "Produit", "Succursale", "Type", "Ville", "Qté"], 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        c.fill = PatternFill("solid", start_color=NAVY)
        c.alignment = Alignment(horizontal="center")

    row = 2
    for r in resultats:
        for s in r["stores"]:
            for col, val in enumerate([r["code"], r["nom"], s["nom"], s["type"], s["ville"], s["qty"]], 1):
                ws2.cell(row=row, column=col, value=val).font = Font(name="Arial", size=9)
            row += 1

    for col, w in zip("ABCDEF", [14, 38, 30, 18, 20, 8]):
        ws2.column_dimensions[get_column_letter(col)].width = w

    wb.save(filepath)
    print(f"\n✓ Fichier sauvegardé : {filepath}")


def main():
    cache = load_cache()
    print(f"Cache existant : {len(cache)} ID(s) déjà connus\n")
    print(f"{'─'*55}")

    resultats = []
    for i, (code, nom) in enumerate(PRODUITS, 1):
        print(f"[{i:02d}/{len(PRODUITS)}] {code} — {nom[:40]}")

        # Étape 1 : ID interne
        internal_id = get_internal_id(code, cache)
        if not internal_id:
            print(f"       ✗ ID interne introuvable")
            resultats.append({"code": code, "nom": nom, "total": 0,
                               "statut": "erreur_id", "erreur": "ID interne non trouvé", "stores": []})
            time.sleep(DELAI)
            continue

        print(f"       ID interne : {internal_id}", end=" → ")

        # Étape 2 : distribution
        dispo = get_distribution(internal_id)
        if dispo["erreur"]:
            print(f"✗ {dispo['erreur']}")
            statut = "erreur_api"
        else:
            print(f"{dispo['total']} succursale(s)")
            statut = "ok"

        resultats.append({
            "code":   code,
            "nom":    nom,
            "total":  dispo["total"],
            "statut": statut,
            "erreur": dispo["erreur"],
            "stores": dispo["stores"],
        })

        time.sleep(DELAI)

    date_str = datetime.today().strftime("%Y%m%d")
    output   = f"Distribution_SAQ_{date_str}.xlsx"
    build_excel(resultats, output)

    ok      = sum(1 for r in resultats if r["statut"] == "ok")
    erreurs = sum(1 for r in resultats if r["statut"] != "ok")
    print(f"\nRécap : {ok} produits OK · {erreurs} erreur(s)")


if __name__ == "__main__":
    main()
