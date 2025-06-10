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
        page.wait_for_selector('xpath=//label[span[contains(text(), "De 1 à 2 kg")]]', timeout=10000)
        page.locator('xpath=//label[span[contains(text(), "De 1 à 2 kg")]]').click()
    except PlaywrightTimeoutError as e:
        logging.error(f"Le filtre 'De 1 à 2 kg' n'a pas été trouvé : {e}")
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

    if "1" not in nb_filtre or "De 1 à 2 kg" not in categorie:
        logging.error("Le nombre de filtres ou la catégorie n'est pas le bon")
        raise ValueError

    logging.info(f"Il y a {nb_filtre}, {nb_produit} produits, et la catégorie est : {categorie}")

    return page

def loop_item(page):
    #Itentifier la page actuelle et le nombre d'élément

    global item
    items = page.locator('xpath=//li[@data-sku]').all()
    current_page = page.locator('xpath=//span[@class="pgActual"]').text_content().strip()
    nb_filtre = page.locator('xpath=//span[contains(@class, "FilterTitle")]').text_content().strip()

    logging.info(f" il y {nb_filtre} sur la page {current_page} qui contient {len(items)} éléments.")

    # Extraire les données
    for item in items:
        if item.locator('xpath =. //a[@data-link]').count() > 0:
            link = item.locator('xpath =. //a[@data-link]').get_attribute('href')
    else:
        link = None

    if item.locator('xpath =. //h2').count() > 0:
        titre = item.locator('xpath =. //h2').text_content.strip()
    else:
        titre = None

    if item.locator('xpath =. //span[contains(@class, "priceColor")]').count() > 0:
            prix_actuel = item.locator('xpath =. //span[contains(@class, "priceColor")]').text_content.strip()
    else:
        prix_actuel = None

    if item.locator('xpath =. //s').count() > 0:
        ancien_prix = item.locator('xpath =. //s').text_content.strip()
    else:
        ancien_prix = None

def main():
    url = "https://www.cdiscount.com/informatique/r-pc+portable+14+pouces.html#_his_"
    SBR_WS_CDP = "wss://brd-customer-hl_c337fd9e-zone-scraping_browser1:oyls3zs0xaeg@brd.superproxy.io:9222"
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(SBR_WS_CDP)
        page = browser.new_page()

        # changer les filtres
        page_criterias = change_criterias(page, url)

        # extraire les donnéees de la première page
        loop_item(page_criterias)


        browser.close()

if __name__ == "__main__":
    main()