import csv
import json
import re
import statistics
import time
import unicodedata
from collections import defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "donnees"
OUTPUT = ROOT / "resultats"
SHOPS = ("Marché Lagune", "Baobab Shop", "Comptoir Ivoire")
HISTORY_SHOPS = {
    "marche lagune": "Marché Lagune",
    "baobab shop": "Baobab Shop",
    "comptoir ivoire": "Comptoir Ivoire",
}

ALIASES = (
    ("P01", ("galaxy a15",)),
    ("P02", ("galaxy a25", "samsung a25")),
    ("P03", ("galaxy a55",)),
    ("P04", ("spark 20",)),
    ("P05", ("camon 30",)),
    ("P06", ("hot 40i",)),
    ("P07", ("note 40",)),
    ("P08", ("itel a70",)),
    ("P09", ("iphone 13",)),
    ("P10", ("redmi note 13",)),
    ("P11", ("hp 250 g9", "hp laptop 250 g9")),
    ("P12", ("ideapad 3",)),
    ("P13", ("inspiron 15",)),
    ("P14", ("sandisk",)),
    ("P15", ("seagate",)),
    ("P16", ("m185",)),
    ("P17", ("deskjet 2720",)),
    ("P18", ("archer c6",)),
    ("P19", ("samsung 24", "24 monitor")),
    ("P20", ("hisense 205", "hisense fridge 205")),
    ("P21", ("nasco 200",)),
    ("P22", ("samsung 43",)),
    ("P23", ("hisense 32",)),
    ("P24", ("lg 1cv", "lg air conditioner 1hp")),
    ("P25", ("binatone 16",)),
    ("P26", ("moulinex",)),
    ("P27", ("philips",)),
    ("P28", ("samsung 23", "samsung microwave 23")),
    ("P29", ("bouilloire", "electric kettle")),
    ("P30", ("hisense 1.5", "hisense air conditioner 1.5")),
)


def normalize(value):
    value = unicodedata.normalize("NFKD", str(value or "").casefold())
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = value.replace("\ufffd", " ")
    return " ".join(re.sub(r"[^a-z0-9]+", " ", value).split())


def repair_mojibake(value):
    if "Ã" not in value and "Â" not in value:
        return value, False
    try:
        repaired = value.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return value, False
    return repaired, repaired != value


def identify_product(label):
    text = normalize(label)
    for product_id, aliases in ALIASES:
        if any(normalize(alias) in text for alias in aliases):
            return product_id
    return None


def parse_price(value):
    digits = re.sub(r"\D", "", str(value or ""))
    return int(digits) if digits else None


def log_issue(journal, source, item, issue, action, details=""):
    journal.append({
        "source": source,
        "article": item,
        "probleme": issue,
        "traitement": action,
        "details": details,
    })


def load_catalogue(journal):
    catalogue = {}
    with (DATA / "catalogue_reference.csv").open(encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            cleaned = {}
            for field in ("nom", "categorie", "marque"):
                cleaned[field], repaired = repair_mojibake(row[field])
                if repaired:
                    log_issue(
                        journal, "Catalogue", row["produit_id"], "Texte mal encodé",
                        "Texte réparé", f"{field}: {row[field]}",
                    )
            catalogue[row["produit_id"]] = {
                "nom": cleaned["nom"],
                "categorie": cleaned["categorie"],
                "marque": cleaned["marque"],
            }
    return catalogue


def make_offer(source, product_id, label, price, available, journal, details=""):
    if not product_id:
        log_issue(journal, source, label, "Article non rapproché", "Exclu")
        return None
    if price is None or price <= 0:
        log_issue(journal, source, label, "Prix absent ou inexploitable", "Exclu")
        return None
    if not available:
        log_issue(journal, source, label, "Article indisponible", "Exclu", details)
        return None
    return {
        "source": source,
        "produit_id": product_id,
        "prix": price,
        "article": label,
    }


def load_offers(journal):
    offers = []
    seen = set()

    with (DATA / "marche_lagune.csv").open(encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file, delimiter=";"):
            label = row["designation"].strip()
            product_id = identify_product(label)
            key = ("Marché Lagune", product_id)
            if product_id and key in seen:
                log_issue(journal, "Marché Lagune", label, "Doublon", "Deuxième ligne ignorée", row["reference"])
                continue
            price = parse_price(row["prix"])
            available = row["disponibilite"].strip().casefold() != "rupture"
            offer = make_offer(
                "Marché Lagune", product_id, label, price, available, journal,
                row["disponibilite"].strip(),
            )
            if offer:
                offers.append(offer)
            if product_id:
                seen.add(key)

    baobab = json.loads((DATA / "baobab_shop.json").read_text(encoding="utf-8"))
    for item in baobab["items"]:
        label = item["title"]
        product_id = identify_product(label)
        key = ("Baobab Shop", product_id)
        if product_id and key in seen:
            log_issue(journal, "Baobab Shop", label, "Doublon", "Deuxième ligne ignorée", item["sku"])
            continue
        pricing = item.get("pricing", {})
        if pricing.get("currency") != "XOF":
            log_issue(journal, "Baobab Shop", label, "Devise inattendue", "Prix exclu", str(pricing.get("currency")))
            price = None
        else:
            amount = pricing.get("amount")
            price = amount + (pricing.get("delivery_fee") or 0) if isinstance(amount, (int, float)) else None
        offer = make_offer(
            "Baobab Shop", product_id, label, price, item.get("stock", 0) > 0, journal,
            "stock nul" if item.get("stock", 0) <= 0 else "",
        )
        if offer:
            offers.append(offer)
        if product_id:
            seen.add(key)

    sheet = load_workbook(DATA / "comptoir_ivoire.xlsx", data_only=True).active
    for row in sheet.iter_rows(min_row=4, values_only=True):
        code, label, _, catalogue_price, discount, stock, _ = row
        if not label:
            continue
        if "\ufffd" in label:
            log_issue(
                journal, "Comptoir Ivoire", label, "Caractère remplacé par le décodage du classeur",
                "Rapprochement conservé à partir de la marque et du modèle",
            )
        product_id = identify_product(label)
        key = ("Comptoir Ivoire", product_id)
        if product_id and key in seen:
            log_issue(journal, "Comptoir Ivoire", label, "Doublon", "Deuxième ligne ignorée", str(code))
            continue
        price = parse_price(catalogue_price)
        if price is not None:
            price = int((Decimal(price) * (Decimal(100) - Decimal(discount or 0)) / Decimal(100))
                        .quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        offer = make_offer(
            "Comptoir Ivoire", product_id, label, price, (stock or 0) > 0, journal,
            "quantité en rayon nulle" if (stock or 0) <= 0 else "",
        )
        if offer:
            offers.append(offer)
        if product_id:
            seen.add(key)

    # Quarantine obvious tenfold-scale errors instead of treating them as market prices.
    by_product = defaultdict(list)
    for offer in offers:
        by_product[offer["produit_id"]].append(offer)
    excluded = set()
    for product_id, product_offers in by_product.items():
        for offer in product_offers:
            other_prices = [other["prix"] for other in product_offers if other["source"] != offer["source"]]
            if other_prices and offer["prix"] > 3 * statistics.median(other_prices):
                excluded.add(id(offer))
                log_issue(
                    journal, offer["source"], offer["article"],
                    "Prix aberrant (> 3 fois la médiane des autres enseignes)",
                    "Exclu du comparatif", f"{offer['prix']} FCFA; produit {product_id}",
                )
    return [offer for offer in offers if id(offer) not in excluded]


def write_csv(path, rows, columns):
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def build_comparison(catalogue, offers):
    prices = defaultdict(dict)
    for offer in offers:
        prices[offer["produit_id"]][offer["source"]] = offer["prix"]

    rows = []
    for product_id, product in catalogue.items():
        shop_prices = prices.get(product_id, {})
        if len(shop_prices) < 2:
            continue
        low = min(shop_prices.values())
        high = max(shop_prices.values())
        cheapest = " / ".join(shop for shop in SHOPS if shop_prices.get(shop) == low)
        row = {
            "produit_id": product_id,
            "produit": product["nom"],
            "moins_cher": cheapest,
            "ecart_fcfa": high - low,
            "ecart_pct_vs_moins_cher": round((high - low) / low * 100, 2),
        }
        row.update({shop: shop_prices.get(shop, "") for shop in SHOPS})
        rows.append(row)
    return rows, prices


def load_history(journal):
    indexed = defaultdict(list)
    with (DATA / "historique_prix.csv").open(encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file):
            source = HISTORY_SHOPS.get(normalize(row["source"]))
            price = parse_price(row["prix_fcfa"])
            try:
                day = date.fromisoformat(row["date"])
            except ValueError:
                day = None
            if source is None or price is None or day is None or row["produit_id"] not in {
                f"P{i:02}" for i in range(1, 31)
            }:
                log_issue(journal, "Historique", row.get("produit_id", ""), "Ligne invalide", "Ligne ignorée", str(row))
                continue
            indexed[(source, row["produit_id"])].append((day, price))
    return indexed


def history_summary(indexed):
    start = date(2025, 11, 24)
    end = date(2025, 11, 30)
    rows = []
    for shop in SHOPS:
        annual = []
        black_friday = []
        products = sorted(product_id for source, product_id in indexed if source == shop)
        for product_id in products:
            values = indexed[(shop, product_id)]
            annual.append(statistics.mean(price for _, price in values))
            week = [price for day, price in values if start <= day <= end]
            if week:
                black_friday.append(statistics.mean(week))
        annual_mean = statistics.mean(annual) if annual else 0
        bf_mean = statistics.mean(black_friday) if black_friday else 0
        rows.append({
            "enseigne": shop,
            "produits_avec_historique": len(annual),
            "produits_semaine_black_friday": len(black_friday),
            "moyenne_annuelle_fcfa": round(annual_mean),
            "moyenne_24_30_nov_fcfa": round(bf_mean),
            "ecart_black_friday_fcfa": round(bf_mean - annual_mean),
            "ecart_black_friday_pct": round((bf_mean / annual_mean - 1) * 100, 2) if annual_mean else 0,
        })
    return rows


def summarize_comparison(comparison, prices):
    winners = defaultdict(int)
    gaps = defaultdict(list)
    for row in comparison:
        product_id = row["produit_id"]
        shop_prices = {shop: prices[product_id][shop] for shop in SHOPS if shop in prices[product_id]}
        low = min(shop_prices.values())
        for shop, price in shop_prices.items():
            if price == low:
                winners[shop] += 1
            gaps[shop].append((price - low, (price - low) / low * 100))
    return winners, gaps


def menu():
    print("\n" + "=" * 58)
    print("       ANALYSE COMPARATIVE DE MARCHE")
    print("=" * 58)
    print("1. Charger les sources")
    print("2. Nettoyer les données et consulter le journal")
    print("3. Rapprocher les articles avec le catalogue")
    print("4. Comparer les prix et exporter le CSV")
    print("5. Afficher la synthèse")
    print("6. Calculer le panier de rentrée")
    print("7. Comparer l'historique et la semaine du Black Friday")
    print("8. Mesurer les performances et expliquer la complexité")
    print("0. Quitter")


def load_current_data():
    journal = []
    catalogue = load_catalogue(journal)
    offers = load_offers(journal)
    comparison, prices = build_comparison(catalogue, offers)
    return {
        "catalogue": catalogue,
        "offers": offers,
        "comparison": comparison,
        "prices": prices,
        "journal": journal,
    }


def show_sources(data):
    counts = defaultdict(int)
    for offer in data["offers"]:
        counts[offer["source"]] += 1
    print(f"\nCatalogue de référence chargé : {len(data['catalogue'])} produits.")
    print("Offres exploitables chargées dans la structure commune :")
    for shop in SHOPS:
        print(f"- {shop}: {counts[shop]} offres")
    print("Formats lus : CSV (Marché Lagune), JSON (Baobab Shop), Excel (Comptoir Ivoire).")


def show_cleaning_log(data):
    path = OUTPUT / "journal_nettoyage.csv"
    write_csv(path, data["journal"], ["source", "article", "probleme", "traitement", "details"])
    print(f"\nAnomalies et traitements consignés : {len(data['journal'])}")
    if not data["journal"]:
        print("Aucune anomalie détectée.")
    else:
        for issue in data["journal"]:
            print(f"- [{issue['source']}] {issue['article']}: {issue['probleme']} -> {issue['traitement']}")
    print(f"Journal exporté : {path}")


def show_matching(data):
    valid = defaultdict(dict)
    for offer in data["offers"]:
        valid[offer["produit_id"]][offer["source"]] = offer
    print("\nRAPPROCHEMENT AU CATALOGUE")
    print(f"{'ID':<5} {'Produit':<43} Enseignes avec offre exploitable")
    print("-" * 100)
    for product_id, product in data["catalogue"].items():
        shops = ", ".join(shop for shop in SHOPS if shop in valid[product_id]) or "Aucune"
        print(f"{product_id:<5} {product['nom']:<43} {shops}")
    unmatched = [row for row in data["journal"] if row["probleme"] == "Article non rapproché"]
    print(f"\nLignes source non rapprochées : {len(unmatched)}")


def show_comparison(data):
    path = OUTPUT / "comparatif.csv"
    columns = ["produit_id", "produit", *SHOPS, "moins_cher", "ecart_fcfa", "ecart_pct_vs_moins_cher"]
    write_csv(path, data["comparison"], columns)
    print(f"\nComparaison des {len(data['comparison'])} produits disponibles dans au moins deux enseignes :")
    for row in data["comparison"]:
        prices = " | ".join(
            f"{shop}: {row[shop]:,} FCFA" for shop in SHOPS if row.get(shop) != ""
        )
        print(f"- {row['produit_id']} {row['produit']}: {prices}")
        print(f"  Moins cher : {row['moins_cher']} | Écart : {row['ecart_fcfa']:,} FCFA "
              f"({row['ecart_pct_vs_moins_cher']:.2f} %)")
    print(f"Comparatif exporté : {path}")


def show_synthesis(data):
    comparison = data["comparison"]
    winners, gaps = summarize_comparison(comparison, data["prices"])
    print(f"\nProduits comparables : {len(comparison)}")
    print("Nombre de fois où l'enseigne est la moins chère (égalité comptée pour chaque enseigne) :")
    for shop in SHOPS:
        print(f"- {shop}: {winners[shop]} fois")
    print("\nÉcart moyen par rapport au prix le plus bas (produits disponibles chez l'enseigne) :")
    for shop in SHOPS:
        values = gaps[shop]
        avg_fcfa = statistics.mean(value[0] for value in values) if values else 0
        avg_pct = statistics.mean(value[1] for value in values) if values else 0
        print(f"- {shop}: {avg_fcfa:,.0f} FCFA ({avg_pct:.2f} %), {len(values)} produits")
    top = sorted(comparison, key=lambda row: row["ecart_fcfa"], reverse=True)[:3]
    print("\nTrois produits où comparer fait le plus économiser :")
    for row in top:
        print(f"- {row['produit_id']} {row['produit']}: jusqu'à {row['ecart_fcfa']:,} FCFA "
              f"({row['ecart_pct_vs_moins_cher']:.2f} %)")


def show_basket(data):
    basket = {"P04": 1, "P12": 1, "P14": 2, "P15": 1, "P28": 1}
    print("\nPANIER DE RENTRÉE")
    for shop in SHOPS:
        missing = [product_id for product_id, quantity in basket.items()
                   if product_id not in data["prices"] or shop not in data["prices"][product_id]]
        if missing:
            print(f"- {shop}: non calculable; indisponible ou sans prix : {', '.join(missing)}")
        else:
            total = sum(data["prices"][product_id][shop] * quantity for product_id, quantity in basket.items())
            print(f"- {shop}: {total:,} FCFA")
    optimized = 0
    for product_id, quantity in basket.items():
        available = data["prices"].get(product_id, {})
        if not available:
            optimized = None
            break
        optimized += min(available.values()) * quantity
    print(f"- En achetant chaque article au moins cher : {optimized:,} FCFA" if optimized is not None else
          "- Panier optimisé non calculable : article sans offre disponible")


def show_history():
    journal = []
    started = time.perf_counter()
    history = load_history(journal)
    rows = history_summary(history)
    elapsed = time.perf_counter() - started
    path = OUTPUT / "historique_black_friday.csv"
    write_csv(path, rows, ["enseigne", "produits_avec_historique", "produits_semaine_black_friday",
                           "moyenne_annuelle_fcfa", "moyenne_24_30_nov_fcfa",
                           "ecart_black_friday_fcfa", "ecart_black_friday_pct"])
    print("\nHISTORIQUE : moyenne annuelle comparée à la semaine du 24 au 30 novembre 2025")
    print("Moyennes calculées par produit puis entre produits; Comptoir = prix catalogue, sans remise.")
    for row in rows:
        print(f"- {row['enseigne']}: Black Friday {row['moyenne_24_30_nov_fcfa']:,} FCFA vs "
              f"année {row['moyenne_annuelle_fcfa']:,} FCFA "
              f"({row['ecart_black_friday_fcfa']:+,} FCFA, {row['ecart_black_friday_pct']:+.2f} %)")
    print(f"Lecture et analyse : {elapsed:.4f} s. CSV exporté : {path}")


def show_performance():
    journal = []
    started = time.perf_counter()
    history = load_history(journal)
    indexed_at = time.perf_counter()
    # Probe repeated keyed lookups to demonstrate the indexed access pattern.
    lookups = [(key, index % len(values)) for index, (key, values) in enumerate(history.items()) if values]
    for _ in range(1000):
        for key, index in lookups:
            history[key][index]
    elapsed = time.perf_counter() - started
    load_elapsed = indexed_at - started
    row_count = sum(len(values) for values in history.values())
    print(f"\nHistorique traité : {row_count:,} lignes.")
    print(f"Construction de l'index : {load_elapsed:.4f} s.")
    print(f"Indexation et 1 000 séries d'accès : {elapsed:.4f} s.")
    print("Structure : dictionnaire indexé par (enseigne, produit_id), contenant les couples (date, prix).")
    print("Complexité : construction O(n), accès moyen O(1), mémoire O(n);")
    print("un balayage de l'historique par recherche coûterait O(n) par recherche.")


def main():
    OUTPUT.mkdir(exist_ok=True)
    data = None
    actions = {
        "1": lambda: show_sources(data),
        "2": lambda: show_cleaning_log(data),
        "3": lambda: show_matching(data),
        "4": lambda: show_comparison(data),
        "5": lambda: show_synthesis(data),
        "6": lambda: show_basket(data),
        "7": show_history,
        "8": show_performance,
    }
    while True:
        menu()
        choice = input("\nVotre choix (0-8) : ").strip()
        if choice == "0":
            print("Fin du programme.")
            break
        if choice not in actions:
            print("Choix invalide. Saisissez un nombre de 0 à 8.")
            continue
        if choice in {"1", "2", "3", "4", "5", "6"} and data is None:
            print("\nChargement et préparation des données...")
            data = load_current_data()
        actions[choice]()
        input("\nAppuyez sur Entrée pour revenir au menu...")


if __name__ == "__main__":
    main()
