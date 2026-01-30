"""
=============================================================================
GENERATEUR DE LEADS GOOGLE MAPS
=============================================================================
Script pour rechercher des entreprises sur Google Maps, analyser leurs sites
web et exporter les données dans un fichier Excel.

Usage: python lead_generator.py
=============================================================================
"""

import os
import re
import googlemaps
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from dotenv import load_dotenv
from datetime import datetime

# Charger les variables d'environnement
load_dotenv()

# =============================================================================
# CONFIGURATION
# =============================================================================

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
TIMEOUT_REQUESTS = 10  # Timeout pour les requêtes HTTP


# =============================================================================
# FONCTIONS PRINCIPALES
# =============================================================================

def rechercher_entreprises(activite: str, localisation: str, rayon_km: int = 10) -> list:
    """
    Recherche des entreprises sur Google Maps.

    Args:
        activite: Type d'entreprise (ex: "restaurant", "plombier")
        localisation: Ville ou adresse (ex: "Paris, France")
        rayon_km: Rayon de recherche en kilomètres

    Returns:
        Liste des entreprises trouvées avec leurs infos
    """
    print(f"\n🔍 Recherche: '{activite}' à '{localisation}' (rayon: {rayon_km}km)")

    # Initialiser le client Google Maps
    gmaps = googlemaps.Client(key=GOOGLE_API_KEY)

    # Géocoder la localisation pour obtenir les coordonnées
    geocode = gmaps.geocode(localisation)
    if not geocode:
        print(f"❌ Localisation '{localisation}' non trouvée")
        return []

    lat = geocode[0]["geometry"]["location"]["lat"]
    lng = geocode[0]["geometry"]["location"]["lng"]
    print(f"📍 Coordonnées: {lat}, {lng}")

    # Rechercher les entreprises
    resultats = []
    next_page_token = None

    while True:
        # Requête Places API
        if next_page_token:
            import time
            time.sleep(2)  # Attendre avant d'utiliser le token
            response = gmaps.places_nearby(
                location=(lat, lng),
                radius=rayon_km * 1000,
                keyword=activite,
                page_token=next_page_token
            )
        else:
            response = gmaps.places_nearby(
                location=(lat, lng),
                radius=rayon_km * 1000,
                keyword=activite
            )

        # Traiter les résultats
        for place in response.get("results", []):
            # Obtenir les détails complets
            details = gmaps.place(place["place_id"], fields=[
                "name", "formatted_address", "formatted_phone_number",
                "website", "rating", "user_ratings_total", "types"
            ])["result"]

            entreprise = {
                "nom": details.get("name", "N/A"),
                "adresse": details.get("formatted_address", "N/A"),
                "telephone": details.get("formatted_phone_number", "N/A"),
                "site_web": details.get("website", None),
                "note": details.get("rating", "N/A"),
                "nb_avis": details.get("user_ratings_total", 0),
                "types": ", ".join(details.get("types", []))
            }

            resultats.append(entreprise)
            print(f"  ✓ {entreprise['nom']}")

        # Pagination
        next_page_token = response.get("next_page_token")
        if not next_page_token:
            break

    print(f"\n✅ {len(resultats)} entreprises trouvées")
    return resultats


def analyser_site_web(url: str) -> dict:
    """
    Analyse un site web et identifie ses points forts et faibles.

    Args:
        url: URL du site à analyser

    Returns:
        Dictionnaire avec points_forts et points_faibles
    """
    print(f"  🌐 Analyse du site: {url}")

    points_forts = []
    points_faibles = []

    try:
        # Récupérer le contenu du site
        headers = {"User-Agent": "Mozilla/5.0 (compatible; LeadBot/1.0)"}
        response = requests.get(url, headers=headers, timeout=TIMEOUT_REQUESTS)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # === ANALYSE SEO ===

        # 1. Balise title
        title = soup.find("title")
        if title and len(title.text.strip()) > 10:
            points_forts.append("Titre de page bien défini")
        else:
            points_faibles.append("Titre de page manquant ou trop court")

        # 2. Meta description
        meta_desc = soup.find("meta", attrs={"name": "description"})
        if meta_desc and meta_desc.get("content", ""):
            points_forts.append("Meta description présente")
        else:
            points_faibles.append("Meta description manquante")

        # 3. Balises H1
        h1_tags = soup.find_all("h1")
        if len(h1_tags) == 1:
            points_forts.append("Une seule balise H1 (bonne pratique)")
        elif len(h1_tags) == 0:
            points_faibles.append("Aucune balise H1")
        else:
            points_faibles.append(f"Plusieurs balises H1 ({len(h1_tags)})")

        # 4. Images avec alt
        images = soup.find_all("img")
        images_avec_alt = [img for img in images if img.get("alt")]
        if images:
            ratio_alt = len(images_avec_alt) / len(images)
            if ratio_alt >= 0.8:
                points_forts.append("Images bien optimisées (attributs alt)")
            elif ratio_alt < 0.5:
                points_faibles.append("Beaucoup d'images sans attribut alt")

        # === ANALYSE CONTENU ===

        # 5. Liens sociaux
        social_patterns = ["facebook", "instagram", "twitter", "linkedin", "youtube"]
        liens = [a.get("href", "") for a in soup.find_all("a", href=True)]
        liens_sociaux = [l for l in liens if any(s in l.lower() for s in social_patterns)]
        if liens_sociaux:
            points_forts.append(f"Présence sur les réseaux sociaux ({len(liens_sociaux)} liens)")
        else:
            points_faibles.append("Aucun lien vers les réseaux sociaux")

        # 6. Formulaire de contact
        forms = soup.find_all("form")
        if forms:
            points_forts.append("Formulaire de contact présent")
        else:
            points_faibles.append("Pas de formulaire de contact visible")

        # 7. HTTPS
        if url.startswith("https://"):
            points_forts.append("Site sécurisé (HTTPS)")
        else:
            points_faibles.append("Site non sécurisé (pas de HTTPS)")

        # 8. Numéro de téléphone visible
        tel_pattern = r"(\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}"
        page_text = soup.get_text()
        if re.search(tel_pattern, page_text):
            points_forts.append("Numéro de téléphone visible")
        else:
            points_faibles.append("Numéro de téléphone non visible")

        # 9. Responsive (viewport meta)
        viewport = soup.find("meta", attrs={"name": "viewport"})
        if viewport:
            points_forts.append("Site adapté mobile (viewport)")
        else:
            points_faibles.append("Site possiblement non adapté mobile")

        # 10. Temps de chargement (approximatif)
        if response.elapsed.total_seconds() < 2:
            points_forts.append("Temps de réponse rapide")
        elif response.elapsed.total_seconds() > 5:
            points_faibles.append("Temps de réponse lent")

    except requests.RequestException as e:
        points_faibles.append(f"Site inaccessible: {str(e)[:50]}")
    except Exception as e:
        points_faibles.append(f"Erreur d'analyse: {str(e)[:50]}")

    return {
        "points_forts": points_forts,
        "points_faibles": points_faibles
    }


def exporter_excel(entreprises: list, nom_fichier: str = None) -> str:
    """
    Exporte les données des entreprises dans un fichier Excel.

    Args:
        entreprises: Liste des entreprises avec leurs données
        nom_fichier: Nom du fichier (optionnel, généré automatiquement si absent)

    Returns:
        Chemin du fichier créé
    """
    if not nom_fichier:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nom_fichier = f"leads_{timestamp}.xlsx"

    print(f"\n📊 Export vers {nom_fichier}")

    # Créer le classeur Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    # En-têtes
    headers = [
        "Nom", "Adresse", "Téléphone", "Site Web",
        "Note", "Nb Avis", "Points Forts", "Points Faibles"
    ]

    # Style des en-têtes
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill

    # Données
    for row, ent in enumerate(entreprises, 2):
        ws.cell(row=row, column=1, value=ent.get("nom", ""))
        ws.cell(row=row, column=2, value=ent.get("adresse", ""))
        ws.cell(row=row, column=3, value=ent.get("telephone", ""))
        ws.cell(row=row, column=4, value=ent.get("site_web", ""))
        ws.cell(row=row, column=5, value=ent.get("note", ""))
        ws.cell(row=row, column=6, value=ent.get("nb_avis", ""))

        # Points forts et faibles (si site analysé)
        analyse = ent.get("analyse", {})
        points_forts = "\n".join(analyse.get("points_forts", []))
        points_faibles = "\n".join(analyse.get("points_faibles", []))

        ws.cell(row=row, column=7, value=points_forts)
        ws.cell(row=row, column=8, value=points_faibles)

    # Ajuster la largeur des colonnes
    largeurs = [30, 50, 20, 40, 10, 10, 40, 40]
    for col, largeur in enumerate(largeurs, 1):
        ws.column_dimensions[chr(64 + col)].width = largeur

    # Sauvegarder
    wb.save(nom_fichier)
    print(f"✅ Fichier créé: {nom_fichier}")

    return nom_fichier


def generer_leads(activite: str, localisation: str, rayon_km: int = 10) -> str:
    """
    Fonction principale: génère des leads et les exporte en Excel.

    Args:
        activite: Type d'entreprise recherché
        localisation: Zone géographique
        rayon_km: Rayon de recherche

    Returns:
        Chemin du fichier Excel généré
    """
    print("=" * 60)
    print("GENERATEUR DE LEADS GOOGLE MAPS")
    print("=" * 60)

    # 1. Rechercher les entreprises
    entreprises = rechercher_entreprises(activite, localisation, rayon_km)

    if not entreprises:
        print("❌ Aucune entreprise trouvée")
        return None

    # 2. Analyser les sites web
    print("\n🔬 Analyse des sites web...")
    for ent in entreprises:
        if ent.get("site_web"):
            ent["analyse"] = analyser_site_web(ent["site_web"])
        else:
            ent["analyse"] = {"points_forts": [], "points_faibles": ["Pas de site web"]}
            print(f"  ⚠️ {ent['nom']}: Pas de site web")

    # 3. Exporter en Excel
    fichier = exporter_excel(entreprises)

    print("\n" + "=" * 60)
    print("TERMINÉ!")
    print("=" * 60)

    return fichier


# =============================================================================
# POINT D'ENTREE
# =============================================================================

if __name__ == "__main__":
    # Vérifier la clé API
    if not GOOGLE_API_KEY:
        print("❌ ERREUR: Clé API Google manquante!")
        print("   Créez un fichier .env avec: GOOGLE_API_KEY=votre_clé")
        exit(1)

    # Exemple d'utilisation
    print("\n" + "=" * 60)
    print("CONFIGURATION DE LA RECHERCHE")
    print("=" * 60)

    # Demander les paramètres à l'utilisateur
    activite = input("\n📌 Type d'activité (ex: restaurant, plombier): ").strip()
    localisation = input("📍 Localisation (ex: Paris, France): ").strip()
    rayon = input("📏 Rayon en km (défaut: 10): ").strip()

    rayon_km = int(rayon) if rayon.isdigit() else 10

    # Lancer la génération
    generer_leads(activite, localisation, rayon_km)
