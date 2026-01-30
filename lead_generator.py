"""
=============================================================================
GENERATEUR DE LEADS GOOGLE MAPS (Scraping)
=============================================================================
Recherche des entreprises sur Google Maps, analyse leurs sites web
et exporte les données dans un fichier Excel.

Usage:
    pip install -r requirements.txt
    playwright install chromium
    python lead_generator.py
=============================================================================
"""

import re
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from urllib.parse import quote

# =============================================================================
# CONFIGURATION
# =============================================================================

DEBUG = False  # Mettre True pour voir le navigateur
SCROLL_PAUSE = 2  # Pause entre chaque scroll (secondes)
MAX_SCROLLS = 20  # Nombre max de scrolls (plus = plus de résultats)


# =============================================================================
# SCRAPING GOOGLE MAPS
# =============================================================================

def rechercher_entreprises(activite: str, localisation: str, max_resultats: int = 100) -> list:
    """
    Scrape Google Maps pour trouver des entreprises.

    Args:
        activite: Type d'entreprise (ex: "plombier")
        localisation: Zone géographique (ex: "Lyon")
        max_resultats: Nombre max de résultats à récupérer

    Returns:
        Liste des entreprises avec leurs infos
    """
    print(f"\n🔍 Recherche: '{activite}' à '{localisation}'")

    # Construire l'URL de recherche
    query = quote(f"{activite} {localisation}")
    url = f"https://www.google.com/maps/search/{query}"

    entreprises = []

    with sync_playwright() as p:
        # Lancer le navigateur (DEBUG=True pour voir ce qui se passe)
        browser = p.chromium.launch(headless=not DEBUG)
        page = browser.new_page()

        print(f"🌐 Chargement de Google Maps...")
        page.goto(url, timeout=60000)

        # Accepter les cookies RGPD (plusieurs variantes FR/EN)
        cookies_buttons = [
            "Tout accepter",
            "Accept all",
            "Accepter tout",
            "J'accepte",
            "Agree",
        ]
        for btn_text in cookies_buttons:
            try:
                page.click(f"button:has-text('{btn_text}')", timeout=2000)
                print(f"✓ Cookies acceptés")
                page.wait_for_timeout(1000)
                break
            except:
                continue

        # Attendre que les résultats chargent (timeout augmenté)
        print(f"⏳ Attente du chargement des résultats...")
        try:
            page.wait_for_selector("div[role='feed']", timeout=30000)
        except:
            # Si toujours pas de feed, prendre un screenshot pour debug
            print(f"❌ Impossible de charger les résultats")
            page.screenshot(path="debug_screenshot.png")
            print(f"📸 Screenshot sauvé: debug_screenshot.png")
            browser.close()
            return []

        print(f"📜 Scroll pour charger plus de résultats...")

        # Scroller pour charger plus de résultats
        feed = page.query_selector("div[role='feed']")
        previous_count = 0

        for i in range(MAX_SCROLLS):
            # Scroller dans le feed
            feed.evaluate("el => el.scrollTop = el.scrollHeight")
            page.wait_for_timeout(SCROLL_PAUSE * 1000)

            # Compter les résultats actuels
            items = page.query_selector_all("div[role='feed'] > div > div[jsaction]")
            current_count = len(items)

            print(f"  Scroll {i+1}/{MAX_SCROLLS} - {current_count} résultats")

            # Arrêter si plus de nouveaux résultats ou max atteint
            if current_count >= max_resultats or current_count == previous_count:
                break
            previous_count = current_count

        # Extraire les données de chaque résultat
        print(f"\n📊 Extraction des données...")
        items = page.query_selector_all("div[role='feed'] > div > div[jsaction]")

        for idx, item in enumerate(items[:max_resultats]):
            try:
                # Cliquer sur l'élément pour voir les détails
                item.click()
                page.wait_for_timeout(1500)

                # Extraire les infos du panneau latéral
                entreprise = extraire_details(page)

                if entreprise and entreprise.get("nom"):
                    entreprises.append(entreprise)
                    print(f"  ✓ {entreprise['nom']}")

            except Exception as e:
                continue  # Passer au suivant si erreur

        browser.close()

    print(f"\n✅ {len(entreprises)} entreprises extraites")
    return entreprises


def extraire_details(page) -> dict:
    """
    Extrait les détails d'une entreprise depuis le panneau Google Maps.

    Args:
        page: Page Playwright

    Returns:
        Dictionnaire avec les infos de l'entreprise
    """
    entreprise = {
        "nom": "",
        "adresse": "",
        "telephone": "",
        "site_web": "",
        "note": "",
        "nb_avis": "",
        "horaires": "",
        "categorie": ""
    }

    try:
        # Nom
        nom_el = page.query_selector("h1")
        if nom_el:
            entreprise["nom"] = nom_el.inner_text().strip()

        # Note et avis
        note_el = page.query_selector("div.F7nice span[aria-hidden='true']")
        if note_el:
            entreprise["note"] = note_el.inner_text().strip()

        avis_el = page.query_selector("div.F7nice span[aria-label*='avis']")
        if avis_el:
            avis_text = avis_el.get_attribute("aria-label")
            nb = re.search(r"(\d+)", avis_text.replace(" ", ""))
            if nb:
                entreprise["nb_avis"] = nb.group(1)

        # Catégorie
        cat_el = page.query_selector("button[jsaction*='category']")
        if cat_el:
            entreprise["categorie"] = cat_el.inner_text().strip()

        # Adresse, téléphone, site web via les boutons d'action
        buttons = page.query_selector_all("button[data-item-id]")
        for btn in buttons:
            item_id = btn.get_attribute("data-item-id") or ""
            aria = btn.get_attribute("aria-label") or ""

            if "address" in item_id:
                entreprise["adresse"] = aria.replace("Adresse:", "").strip()

            elif "phone" in item_id:
                entreprise["telephone"] = aria.replace("Téléphone:", "").strip()

        # Site web (lien séparé)
        site_el = page.query_selector("a[data-item-id='authority']")
        if site_el:
            entreprise["site_web"] = site_el.get_attribute("href")

    except Exception as e:
        pass  # Retourner ce qu'on a pu extraire

    return entreprise


# =============================================================================
# ANALYSE DE SITE WEB
# =============================================================================

def analyser_site_web(url: str) -> dict:
    """
    Analyse un site web et identifie ses points forts et faibles.

    Args:
        url: URL du site à analyser

    Returns:
        Dictionnaire avec points_forts et points_faibles
    """
    print(f"  🌐 Analyse: {url[:50]}...")

    points_forts = []
    points_faibles = []

    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        response = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # === SEO ===

        # Titre
        title = soup.find("title")
        if title and len(title.text.strip()) > 10:
            points_forts.append("Titre de page présent")
        else:
            points_faibles.append("Titre manquant ou trop court")

        # Meta description
        meta_desc = soup.find("meta", attrs={"name": "description"})
        if meta_desc and meta_desc.get("content"):
            points_forts.append("Meta description présente")
        else:
            points_faibles.append("Meta description manquante")

        # H1
        h1_tags = soup.find_all("h1")
        if len(h1_tags) == 1:
            points_forts.append("Structure H1 correcte")
        elif len(h1_tags) == 0:
            points_faibles.append("Pas de balise H1")
        else:
            points_faibles.append(f"Trop de H1 ({len(h1_tags)})")

        # Images alt
        images = soup.find_all("img")
        if images:
            avec_alt = len([i for i in images if i.get("alt")])
            ratio = avec_alt / len(images)
            if ratio >= 0.7:
                points_forts.append("Images optimisées (alt)")
            elif ratio < 0.3:
                points_faibles.append("Images sans attribut alt")

        # === CONTENU ===

        # Réseaux sociaux
        social = ["facebook", "instagram", "twitter", "linkedin", "youtube"]
        liens = [a.get("href", "") for a in soup.find_all("a", href=True)]
        if any(s in " ".join(liens).lower() for s in social):
            points_forts.append("Réseaux sociaux présents")
        else:
            points_faibles.append("Pas de réseaux sociaux")

        # Formulaire
        if soup.find("form"):
            points_forts.append("Formulaire de contact")
        else:
            points_faibles.append("Pas de formulaire visible")

        # === TECHNIQUE ===

        # HTTPS
        if url.startswith("https"):
            points_forts.append("Site sécurisé (HTTPS)")
        else:
            points_faibles.append("Pas de HTTPS")

        # Mobile
        if soup.find("meta", attrs={"name": "viewport"}):
            points_forts.append("Compatible mobile")
        else:
            points_faibles.append("Non optimisé mobile")

        # Vitesse
        if response.elapsed.total_seconds() < 2:
            points_forts.append("Chargement rapide")
        elif response.elapsed.total_seconds() > 5:
            points_faibles.append("Chargement lent")

    except Exception as e:
        points_faibles.append(f"Site inaccessible")

    return {"points_forts": points_forts, "points_faibles": points_faibles}


# =============================================================================
# EXPORT EXCEL
# =============================================================================

def exporter_excel(entreprises: list, nom_fichier: str = None) -> str:
    """
    Exporte les leads dans un fichier Excel.

    Args:
        entreprises: Liste des entreprises
        nom_fichier: Nom du fichier (optionnel)

    Returns:
        Chemin du fichier créé
    """
    if not nom_fichier:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nom_fichier = f"leads_{timestamp}.xlsx"

    print(f"\n📊 Export: {nom_fichier}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    # En-têtes
    headers = ["Nom", "Catégorie", "Adresse", "Téléphone", "Site Web",
               "Note", "Avis", "Points Forts", "Points Faibles"]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill

    # Données
    for row, e in enumerate(entreprises, 2):
        ws.cell(row=row, column=1, value=e.get("nom", ""))
        ws.cell(row=row, column=2, value=e.get("categorie", ""))
        ws.cell(row=row, column=3, value=e.get("adresse", ""))
        ws.cell(row=row, column=4, value=e.get("telephone", ""))
        ws.cell(row=row, column=5, value=e.get("site_web", ""))
        ws.cell(row=row, column=6, value=e.get("note", ""))
        ws.cell(row=row, column=7, value=e.get("nb_avis", ""))

        analyse = e.get("analyse", {})
        ws.cell(row=row, column=8, value="\n".join(analyse.get("points_forts", [])))
        ws.cell(row=row, column=9, value="\n".join(analyse.get("points_faibles", [])))

    # Largeur colonnes
    for col, width in enumerate([25, 20, 40, 15, 35, 8, 10, 35, 35], 1):
        ws.column_dimensions[chr(64 + col)].width = width

    wb.save(nom_fichier)
    print(f"✅ Fichier créé: {nom_fichier}")
    return nom_fichier


# =============================================================================
# FONCTION PRINCIPALE
# =============================================================================

def generer_leads(activite: str, localisation: str, max_resultats: int = 100) -> str:
    """
    Génère des leads et les exporte en Excel.

    Args:
        activite: Type d'entreprise
        localisation: Zone géographique
        max_resultats: Nombre max de leads

    Returns:
        Chemin du fichier Excel
    """
    print("=" * 60)
    print("GENERATEUR DE LEADS GOOGLE MAPS")
    print("=" * 60)

    # 1. Scraper Google Maps
    entreprises = rechercher_entreprises(activite, localisation, max_resultats)

    if not entreprises:
        print("❌ Aucune entreprise trouvée")
        return None

    # 2. Analyser les sites web
    print("\n🔬 Analyse des sites web...")
    for e in entreprises:
        if e.get("site_web"):
            e["analyse"] = analyser_site_web(e["site_web"])
        else:
            e["analyse"] = {"points_forts": [], "points_faibles": ["Pas de site web"]}

    # 3. Exporter
    fichier = exporter_excel(entreprises)

    print("\n" + "=" * 60)
    print("TERMINÉ!")
    print("=" * 60)

    return fichier


# =============================================================================
# POINT D'ENTREE
# =============================================================================

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("CONFIGURATION")
    print("=" * 60)

    activite = input("\n📌 Activité (ex: plombier, restaurant): ").strip()
    localisation = input("📍 Localisation (ex: Lyon, Paris 15): ").strip()
    max_res = input("📊 Nombre max de résultats (défaut: 100): ").strip()

    max_resultats = int(max_res) if max_res.isdigit() else 100

    generer_leads(activite, localisation, max_resultats)
