from pathlib import Path
from RPA.Browser.Selenium import Selenium
from anticaptchaofficial.imagecaptcha import imagecaptcha
from RPA.FileSystem import FileSystem
import time


def anti_captcha(captcha_img: str) -> str:
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key("32de7a9e64b12072de66663439c86b58")

    solver.set_soft_id(0)

    captcha_text = solver.solve_and_return_solution(captcha_img)
    if captcha_text != 0:
        print("captcha text " + captcha_text)
        return captcha_text
    else:
        print("task finished with error " + solver.error_code)
        return False
    
def log_in_and_download_sesi_dr(chrome_browser: Selenium, default_download_directory: str, nfs: list[int]):
    chrome_prefs = {
        "download.default_directory": default_download_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }

    chrome_browser = Selenium()
    chrome_browser.auto_close=False
    url = "https://iss.fortaleza.ce.gov.br/grpfor/login.seam;jsessionid=1GC7+OZ2wq6sBwl0ZemSuXHB.pistol:iss-prod-02?cid=151891"
    chrome_browser.open_chrome_browser(url=url, preferences=chrome_prefs)
    while chrome_browser.does_page_contain_element('login:captchaDecor:captchaLogin'):
        chrome_browser.wait_until_element_is_visible('login:password')
        chrome_browser.input_text('login:username', '31480268372')
        chrome_browser.input_text('login:password', 'senai@2020')
        chrome_browser.capture_element_screenshot('//*[@id="login:captchaDecor"]/img', 'captcha.png')
        chrome_browser.input_text('login:captchaDecor:captchaLogin',anti_captcha("captcha.png"))
        chrome_browser.click_element_if_visible('login:botaoEntrar')