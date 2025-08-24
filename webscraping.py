from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urljoin
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, StaleElementReferenceException

website = 'https://www.carsome.my/buy-car'
path = r'C:\Users\kuokh\OneDrive\Documents\code\chromedriver-win64\chromedriver-win64\chromedriver.exe'
LAST_PAGE = 114 # change if it grows
HEADLESS = True # set False to watch it
POLITE_DELAY_BETWEEN_DETAILS = 0.8
LIGHT_SCROLLS_PER_LISTING_PAGE = 3
OUTPUT_XLSX = "testcars(51-100).xlsx"

overview_keys = [
        "Mileage", "Transmission", "Registration Date",
        "Principal Warranty", "Fuel Type", "Seat"
    ]


service = Service(executable_path=path)
driver = webdriver.Chrome(service=service)
driver.get(website)

def make_driver(headless=True):
    from selenium.webdriver.chrome.options import Options
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

def collect_listing_links_on_page(driver):
    """
    Collects detail URLs from the current listing page.

    Looks for anchors like:
      <a class="mod-b-card__title" href="/buy-car/.../.../.../id">
    """
    # 1) Wait until at least one card-title anchor is in the DOM
    WebDriverWait(driver, 25).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a.mod-b-card__title[href^="/buy-car/"]'))
    )

    # 2) Use JS to read every matching href (more reliable than .find_elements on some SPAs)
    hrefs = driver.execute_script("""
        return Array.from(document.querySelectorAll('a.mod-b-card__title[href^="/buy-car/"]'))
                    .map(a => a.getAttribute('href'));
    """) or []

    # 3) Normalize to absolute URLs and dedupe
    links = { urljoin(website, h.split("?")[0]) for h in hrefs if h and h.startswith("/buy-car/") }

    return links

def dismiss_blocking_modals(driver, timeout=3):
    """
    Dismiss Carsome's 'This car has been ordered' modal by clicking div.dynamic__close if present.
    """
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.dynamic__close"))
        )
        el.click()
        WebDriverWait(driver, 2).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.dynamic__close"))
        )
        return True
    except Exception:
        return False
    
def get_text_or_none(driver, by, value):
    try:
        return driver.find_element(by, value).text
    except NoSuchElementException:
        return None

def extract_overview_from_detail_page(driver, url):

    driver.get(url)

    # Attempt to dismiss the modal, if present
    dismiss_blocking_modals(driver, timeout=1)

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'mod-vehicle-overview')]"))
    )

    car = get_text_or_none(
        driver, By.XPATH, '//*[@id="detail-page"]/div/div/div[2]/div[2]/section[1]/div[2]'
    )

    price = get_text_or_none(
        driver, By.XPATH, '/html/body/div[1]/div/div/div[1]/main/div/div/div/div/div[2]/div[2]/div/section[1]/div/div/span[1]'
    )

    # Wait for main overview elements to load
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "car-value__item-value"))
        )
    except TimeoutException:
        return {}

    overview = {}
    overview["Car"] = car  # Add as the first column
    overview["Price"] = price

    elements = driver.find_elements(By.CLASS_NAME, "car-value__item-value")
    for i, key in enumerate(overview_keys):
        overview[key] = elements[i].text if i < len(elements) else None

    specifications_tab = url + "/specifications"
    driver.get(specifications_tab)

    dismiss_blocking_modals(driver, timeout= 1)

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "item-value"))
        )

        overview["Horse power"] = get_text_or_none(driver, By.XPATH, "//*[@id='app']/div/main/div/div/div[3]/div[1]/div[3]/div/div[2]/div[12]/div/div[2]")
        overview["Torque RPM"] = get_text_or_none(driver, By.XPATH, "//*[@id='app']/div/main/div/div/div[3]/div[1]/div[3]/div/div[2]/div[14]/div/div[2]")
        overview["Cylinders"] = get_text_or_none(driver, By.XPATH, "//*[@id='app']/div/main/div/div/div[3]/div[1]/div[3]/div/div[2]/div[17]/div/div[2]")
        overview["Fuel consumption"] = get_text_or_none(driver, By.XPATH, "//*[@id='app']/div/main/div/div/div[3]/div[1]/div[4]/div/div[2]/div[1]/div/div[2]")

    except TimeoutException:
        pass

    return overview

def accept_cookies_if_any(driver):
    try:
        for txt in ("Accept", "I agree", "OK", "Got it"):
            btn = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, f"//button[contains(., '{txt}')]"))
            )
            btn.click(); time.sleep(0.3); break
    except Exception:
        pass

def infinite_scroll_to_bottom(driver, max_scrolls=10, sleep_each=0.7):
    last_h = driver.execute_script("return document.body.scrollHeight")
    for _ in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(sleep_each)
        new_h = driver.execute_script("return document.body.scrollHeight")
        if new_h == last_h:
            break
        last_h = new_h


def main():
    driver = make_driver(headless=HEADLESS)

    # Build page list: page 1 (no param), then ?pageNo=2..N
    page_urls = [f"{website}?pageNo={i}" for i in range(51, 101)]

    all_detail_urls = set()

    try:
        # -------- 1) Collect all detail links from every listing page --------
        for idx, url in enumerate(page_urls, 1):
            print(f"[Listing] {idx}/{len(page_urls)}: {url}")
            driver.get(url)

            # Wait until card anchors exist, scroll a bit to load more cards, handle cookies
            try:
                WebDriverWait(driver, 25).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'a.mod-b-card__title[href^="/buy-car/"]')
                    )
                )
            except Exception:
                # If anchors didn’t render immediately, try deeper scroll then re-check
                infinite_scroll_to_bottom(driver, max_scrolls=10, sleep_each=0.8)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'a.mod-b-card__title[href^="/buy-car/"]')
                    )
                )

            accept_cookies_if_any(driver)
            infinite_scroll_to_bottom(driver, max_scrolls=LIGHT_SCROLLS_PER_LISTING_PAGE, sleep_each=0.7)

            page_links = collect_listing_links_on_page(driver)
            print(f"  +{len(page_links)} links")
            all_detail_urls.update(page_links)
            time.sleep(0.5)  # be polite between listing pages

        print(f"Total unique detail URLs collected: {len(all_detail_urls)}")

        # -------- 2) Visit each detail URL in a new tab, extract, close --------
        rows = []
        listing_handle = driver.current_window_handle
        for i, url in enumerate(sorted(all_detail_urls), 1):
            print(f"[{i}/{len(all_detail_urls)}] {url}")
            driver.switch_to.new_window("tab")
            try:
                rows.append(extract_overview_from_detail_page(driver, url))
            except Exception as e:
                print(f"  error: {e}")
                rows.append({"url": url, "error": str(e)})
            finally:
                driver.close()
                driver.switch_to.window(listing_handle)
                time.sleep(POLITE_DELAY_BETWEEN_DETAILS)

        # -------- 3) DataFrame first, then write once to Excel --------
        df = pd.DataFrame(rows)
        with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="cars")
        print(f"Saved {len(df)} rows → {OUTPUT_XLSX}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()