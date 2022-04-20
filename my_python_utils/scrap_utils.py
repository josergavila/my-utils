import time

from fp.fp import FreeProxy
from fake_useragent import UserAgent
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

__all__ = (
    "get_chrome_driver",
    "get_firefox_driver",
)

MY_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"


def get_spoofer(country_id: list = ["BR"], rand: bool = True, anonym: bool = True):
    """helper function to retrieve user-agents using free proxy"""
    ua = UserAgent()
    proxy = FreeProxy(country_id=country_id, rand=rand, anonym=anonym).get()
    ip = proxy.split("://")[1]
    return ua.random, ip


def _set_chrome_options(options):
    """helper function to set chrome options"""
    # NOTE: this options are usually enough to avoid detection
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--single-process")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--incognito")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("disable-infobars")
    return options


def get_chrome_driver(use_proxies: bool = False, set_options: bool = True):
    """instantiates and returns a chrome driver"""

    options = webdriver.ChromeOptions()
    if use_proxies:
        userAgent, ip = get_spoofer()
        options.add_argument(f"user-agent={userAgent}")
        options.add_argument(f"--proxy-server={ip}")
    else:
        options.add_argument(f"user-agent={MY_USER_AGENT}")

    if set_options:
        options = _set_chrome_options(options)

    while True:
        try:
            driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
            break
        except FreeProxyException:
            userAgent, ip = get_spoofer()
            options.add_argument(f"user-agent={userAgent}")
            options.add_argument(f"--proxy-server={ip}")

    return driver


def get_firefox_driver():
    """instantiates and returns a firefox driver"""
    options = webdriver.firefox.options.Options()
    options.set_preference(
        "general.useragent.override",
        "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:95.0) Gecko/20100101 Firefox/95.0",
    )
    driver = webdriver.Firefox(options=options)
    return driver


if __name__ == "__main__":

    # a, b = get_spoofer()
    driver = get_chrome_driver()
    url = "https://www.imovelweb.com.br/terrenos-loteamento-condominio-venda-sorocaba-sp.html"
    driver.get(url)
    time.sleep(10)
    driver.close()
