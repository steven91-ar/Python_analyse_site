"""
- fonction pour changer les critères
- fontion pour la loop item
- fontion pour l'extraction des données
- fonction pour exporter sous Excel
- Fonction pour la pagination

"""
import logging

from openpyxl.worksheet import page
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd

logging.basicConfig(level=logging.INFO,
                    filename="cdiscount. log",
                    filemode="w",
                    format="%(asctime)s - %(levelname)s -%(message)s",
                    datefmt="%d-%m-%Y | %H:%M:%S" )

def change_criterias(page, url):
    """
    Appliquer des filtres à partir d'une page PlayWright
    :param page: la page PlayWright
    :param url:
    :return:
    """
    page.goto(url)

    # cliquer sur "tout supprimer"
    try:
        page.wait_for_selector('xpath=//label[span[contains(text(), "De 15")]]', timeout=10000)
        page.locator('xpath=//label[span[contains(text(), "De 15")]]').click()
    except PlaywrightTimeoutError as e:
        logging.error(f"Le filtre 'De 15' n'a pas été trouvé : {e}")
        raise

    try:
        page.wait_for_selector('xpath=//span[contains(@class, "FilterTitle")]', timeout=10000)
        nb_filtre = page.locator('xpath=//span[contains(@class, "FilterTitle")]').text_content().strip()
        nb_filtre = nb_filtre.strip()
    except PlaywrightTimeoutError as e:
        logging.error(f"Impossible de trouver le nombre de filtres : {e}")
        raise

    try:
        categorie = page.locator(
            'xpath=//span[contains(@class, "FilterTitle")]/following-sibling::span[1]').text_content().strip()
        categorie = categorie.strip()
    except PlaywrightTimeoutError as e:
        logging.error(f"La catégorie n'existe pas : {e}")
        raise

    try:
        nb_produit = page.locator('xpath=//h1/following-sibling::span[1]').text_content().strip()
        nb_produit = nb_produit.strip()
    except PlaywrightTimeoutError as e:
        logging.error(f"Le nombre de produits n'existe pas : {e}")
        raise

    if "1" not in nb_filtre or "De 15" not in categorie:
        logging.error("Le nombre de filtres ou la catégorie n'est pas le bon")
        raise ValueError

    logging.info(f"Il y a {nb_filtre}, {nb_produit} produits, et la catégorie est : {categorie}")

    return page

def loop_item(page):
    """
    Extraire les données de chaque produit à partir d'une page Playwright.

    :param page: Objet Playwright représentant la page.
    :return: Liste de dictionnaires contenant les informations des produits.
    """

    items_data = []

    # Identifier la page actuelle et le nombre d'éléments
    items = page.locator('xpath=//li[@data-sku]').all()
    current_page = page.locator('xpath=//span[@class="pgActual"]').text_content().strip()
    nb_filtre = page.locator('xpath=//span[contains(@class, "FilterTitle")]').text_content().strip()

    logging.info(f"Il y a {nb_filtre} sur la page {current_page}, qui contient {len(items)} éléments.")

    # Extraire les données
    for item in items:
        # URL
        if item.locator('xpath=.//a[contains(@class, "o-card__link")]').count() > 0:
            link = item.locator('xpath=.//a[contains(@class, "o-card__link")]').get_attribute('href')
        else:
            link = None

        # Titre
        if item.locator('xpath=.//h4').count() > 0:
            titre = item.locator('xpath=.//h4').text_content().strip()
        else:
            titre = None

        # Prix actuel
        if item.locator('xpath=.//div[contains(@class, "o-card__price")]').count() > 0:
            prix_actuel = item.locator('xpath=.//div[contains(@class, "o-card__price")]').first.text_content().strip()
        else:
            prix_actuel = None

        # Ancien prix
        if item.locator('xpath=.//s').count() > 0:
            ancien_prix = item.locator('xpath=.//s').text_content().strip()
        else:
            ancien_prix = None

        # Mettre les données dans un dictionnaire
        product_details = {
            "url": link,
            "titre": titre,
            "prix": prix_actuel,
            "ancien_prix": ancien_prix,
        }

        items_data.append(product_details)

    return items_data

def export_to_excel(products):
    """
    créer 2 fichier Excel à partir d'une liste d'un produit
    :param products: la liste des produit de cdiscount
    :return:
    """
    #Créer un dataframe avec le prix
    df = pd.DataFrame(products)

    # Modifier la colonne 'prix'
    if "prix" in df.columns:
        df["prix"] = df["prix"].str.replace("€", ",", regex=False) + "€"
    else:
        logging.error(" la colonne 'prix' n'existe pas. ")
        raise KeyError

    # Enregistrer le Dataframe complet sous excel
    try:
        df.to_excel("cdiscount.xlsx", index=False)
        logging.info("Le Dataframe complet a bien été enregistré.")
    except ValueError as e:
        logging.error(f"Le Dataframe complet n'a pas enregistré: {e}")
        raise

    #créer un dataframe avec les produits en promo
    if "ancien_prix" in df.columns:
        df_promo = df.dropna(subset=["ancien_prix"])
    else:
        logging.error(" la colonne 'prix' n'existe pas. ")
        raise KeyError

        # Enregistrer le new Dataframe sous excel
    try:
        df_promo.to_excel("cdiscount_promo.xlsx", index=False)
        logging.info("Le Dataframe en promo a bien été enregistré.")
    except ValueError as e:
        logging.error(f"Le Dataframe promo n'a pas enregistré: {e}")
        raise

def pagination(page):
    """
    Identifier et cliquer sur le bouton 'page suivante'
    :param page:
    :return: True si la page suivante a été ouverte, sinon False
    """
    try:
        page.wait_for_selector('xpath=//li[a[@class="current"]]/following-sibling::li[1]/a', timeout=10000)
        page.locator('xpath=//li[a[@class="current"]]/following-sibling::li[1]/a').click()
        page.wait_for_selector('xpath=//li[@data-sku][1]')
        page.wait_for_timeout(1000)
        return True
    except PlaywrightTimeoutError as e:
        logging.info(f"Nous sommes sur la dernière page : {e}")
        return False


def main():
    url = "https://www.cdiscount.com/informatique/r-pc+portable+14+pouces.html#_his_"
    SBR_WS_CDP = "wss://brd-customer-hl_c337fd9e-zone-scraping_browser1:oyls3zs0xaeg@brd.superproxy.io:9222"
    all_items = []

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(SBR_WS_CDP)
        page = browser.new_page()

        # Changer les filtres
        page_criterias = change_criterias(page, url)

        for i in range(20):
            # Extraire les données de la page courante
            produits = loop_item(page_criterias)
            all_items.extend(produits)
            if not pagination(page_criterias):
                break

        # Exporter les données en Excel
        export_to_excel(all_items)

        browser.close()


if __name__ == "__main__":
    main()