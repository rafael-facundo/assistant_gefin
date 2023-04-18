from pathlib import Path
from RPA.Browser.Selenium import Selenium
from anticaptchaofficial.imagecaptcha import imagecaptcha
from RPA.FileSystem import FileSystem
import time


def anti_captcha(captcha_img: str) -> str:
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key("2ab05e6bf0186f7724d613b158b7365f")

    solver.set_soft_id(0)

    captcha_text = solver.solve_and_return_solution(captcha_img)
    if captcha_text != 0:
        print("captcha text " + captcha_text)
        return captcha_text
    else:
        print("task finished with error " + solver.error_code)
        return False
    

def check_if_download_has_finished(directory_where_files_are: str) -> bool:
    file_system = FileSystem()
    try:
        list_of_file_names = file_system.list_files_in_directory(
            directory_where_files_are)
        for file in list_of_file_names:
            if file.name[-11:] == ".crdownload" or file.name[-4:] == ".tmp":
                return False
            return True
    except FileNotFoundError:
        return False


def temp_filename_for_downloads(downloads_dir: str, nf: int):
    file = FileSystem()
    old_name = f"{downloads_dir}\\relatorio.pdf"
    new_name = f"{downloads_dir}\\nfes_{nf}.pdf"
    while not check_if_download_has_finished(downloads_dir):
        pass
    time.sleep(1)
    file.move_file(old_name, new_name, True)


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
    chrome_browser.click_element_when_visible("alteraInscricaoForm:empresaDataTable:5:linkDocumento")
    chrome_browser.wait_until_element_is_not_visible("alteraInscricaoForm:empresaDataTable:5:linkDocumento")
    for nf in nfs:
        if type(nf) == int :
            chrome_browser.go_to('https://iss.fortaleza.ce.gov.br/grpfor/pages/nfse/consultaNfsePorNumero.seam?cid=212315')
            print(nf)
            chrome_browser.input_text('//*[@id="consultarnfseForm:numNfse"]', nf)
            chrome_browser.click_element('consultarnfseForm:j_id227')
            chrome_browser.click_element_when_visible('//*[@id="consultarnfseForm:dataTable:0:j_id356"]/a')
            chrome_browser.click_element_when_visible('//*[@id="j_id153:panelAcoes"]/tbody/tr/td[1]/input')
            while not check_if_download_has_finished(default_download_directory):
                pass
            temp_filename_for_downloads(default_download_directory, nf)    
        else:
            pass    
    chrome_browser.close_browser()
