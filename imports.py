from RPA.Email.ImapSmtp import ImapSmtp
from RPA.Excel.Files import Files
from RPA.Assistant import Assistant
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path
from RPA.Browser.Selenium import Selenium
import time
from anticaptchaofficial.imagecaptcha import imagecaptcha

#classe para guardar informações de clientes
class ClientInfo:
    name: str
    number: int
    rps: int
    body: str
    email: str


    def __init__(self, name: str, number: int, rps: int, email="none") -> None:
        try:
            self.name = name.replace("/", "").replace(".", "")
        except AttributeError:
            self.name = "sem_nome"
        self.number = number
        self.rps = rps
        try:
            self.email = email
        except:
            self.email = "sem_email"

#função para buscar o número de rps dentro do boleto
def get_number_rps(num_page, directory: str):
    pdf = PDF()
    pdf.open_pdf(directory) 
    number_nf = pdf.find_text("Nro.Documento", num_page, "down") 
    number_nf = number_nf[2].neighbours[0] 
    number_nf = number_nf.lstrip('0') 
    return number_nf 

#função para abrir pdf e salvar páginas já renomeadas
def open_pdf_and_save_pages(directory: str):
    pdf = PDF()
    file = FileSystem() 
    pdf.open_pdf(directory)
    for i in range(1, pdf.get_number_of_pages(), 2):
        # pdf.extract_pages_from_pdf(r"input\\Boleto.pdf",f"output\\boleto_{i}.pdf", i)
        pdf.extract_pages_from_pdf(directory,f"{Path().home()}\\Documents\\Documentos RPA FIEC\\boletos separados\\boleto_{i}.pdf", i)

        # old_name = f"output\\boleto_{i}.pdf"
        old_name = f"{Path().home()}\\Documents\\Documentos RPA FIEC\\boletos separados\\boleto_{i}.pdf"
        
        #new_name = f"output\\output boleto\\RPS {get_number_rps(i)} BOLETO.pdf"
        new_name = f"{Path().home()}\\Documents\\Documentos RPA FIEC\\boletos separados\\RPS {get_number_rps(i, directory)} BOLETO.pdf"
        file.move_file(old_name, new_name) 

#função para enviar emails
def send_email(list_of_client_info: list[ClientInfo], nfs_directory: str, boletos_directory: str):
    for info in list_of_client_info:
        nf = f"{nfs_directory}\\NF {info.number} RPS {info.rps}.pdf"
        boleto = f"{boletos_directory}\\RPS {info.rps} BOLETO.pdf"
        gmail_account = "rfcosta@sfiec.org.br"
        gmail_app_password = "jiitxvjdnglplqxx"
        mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
        mail.authorize(account=gmail_account, password=gmail_app_password)
        mail.send_message(sender=gmail_account,
                            recipients= info.email,
                            subject=f"NF {info.number} RPS {info.rps} - {info.name}",
                            body=f"""Prezado(a) cliente, bom dia. \n\n
Seguem anexos nota fiscal e boleto referente ao serviço solicitado \n
Para informações complementares relacionadas a nota fiscal, entrar em contato com a área de faturamento: 3421.5894 / gefin.iel@sfiec.org.br.\n\n
Por gentileza, confirmar o recebimento e encaminhar ao responsável financeiro de sua empresa.\n\n
att,\n\n
Equipe de faturamento.""",
                            attachments=[nf, boleto])
        
#função para interface visual
def start_dialog(op_list: list[str]):
    dialog = Assistant()
    dialog.add_image(r"Logo-FIEC-02-Color-1024-x-337.jpg")
    dialog.add_drop_down("website_choice",
                         op_list,
                         label="O que deseja fazer?")
    dialog.add_file_input("input_excel_file",
                          "Selecione uma planilha",
                          file_type="xlsx")
    dialog.add_file_input("add_pdf_file", 'Selecione um boleto', file_type="pdf")
    dialog.add_file_input("add_pdf_directory", 'Selecione o diretório das nfs', file_type="pdf")
    dialog.add_submit_buttons(["Executar", "Cancelar"])
    user_input = dialog.run_dialog(90000, title="Assistente - GEFIN", height=550, width=400)

    return user_input

#função para ler planilha
def read_excel(excel_file_path: str):
    excel = Files()
    excel.open_workbook(excel_file_path)
    excel_dict = excel.read_worksheet(header=True)
    list_of_excel = []
    for line in excel_dict:
        list_of_excel.append(ClientInfo(line.get("RAZÃO SOCIAL"), line.get("NOTA FISCAL"), line.get("RPS"), line.get("E-MAILS")))
    return list_of_excel

#função para resolver captcha
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

#função para verificar se downloads foram concluídos
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

#função para renomear temporariamente as notas do iss "relatorio"
def temp_filename_for_downloads(downloads_dir: str, nf: int):
    file = FileSystem()
    old_name = f"{downloads_dir}\\relatorio.pdf"
    new_name = f"{downloads_dir}\\nfes_{nf}.pdf"
    while not check_if_download_has_finished(downloads_dir):
        pass
    time.sleep(1)
    file.move_file(old_name, new_name, True)

#função para renomear de forma definitiva as nfs do sesi dr
def rename_pdf_sesi_dr(directory: str):
    file = FileSystem()
    pdf = PDF()
    list_of_file_names = file.list_files_in_directory(directory)
    for file_adr in list_of_file_names:
        pdf.open_pdf(file_adr)
        rps_matches = pdf.find_text("Número do RPS")
        rps = rps_matches[0].neighbours[0]
        file_str = str(file_adr)
        nf_number = file_str[-8:-4]
        new_name = f"{file_str[:-16]}DR\\NF {nf_number} RPS {rps}.pdf"
        pdf.close_pdf()
        file.move_file(file_adr, new_name)

def rename_pdf_sesi_albano(directory):
    file = FileSystem()
    pdf = PDF()
    list_of_file_names = file.list_files_in_directory(directory)
    for file_adr in list_of_file_names:
        old_name = str(file_adr)
        pdf.open_pdf(file_adr)
        rps_matches = pdf.find_text("Nº do RPS")
        rps = rps_matches[0].neighbours[0]
        nf = old_name[-9:-4]
        new_name = f"{old_name[:-14]}\\NF {nf} RPS {rps}.pdf"
        pdf.close_pdf()
        file.move_file(old_name, new_name)

def rename_pdf_iel(directory):
    file = FileSystem()
    pdf = PDF()
    list_of_file_names = file.list_files_in_directory(directory)
    for file_adr in list_of_file_names:
        old_name = str(file_adr)
        pdf.open_pdf(file_adr) 
        rps_matches = pdf.find_text("Número do RPS")
        rps = rps_matches[0].neighbours[0]
        nf = old_name[-9:-4]
        new_name = f"{old_name[:-15]}\\NF {nf} RPS {rps}.pdf"
        pdf.close_pdf()
        file.move_file(old_name, new_name)
        


#função para baixar as notas do sesi dr
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

# função para baixar as notas do iel
def log_in_and_download_iel(chrome_browser: Selenium, default_download_directory: str, nfs: list[int]):
    chrome_prefs = {
        "download.default_directory": default_download_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }
    cont = 0
    chrome_browser = Selenium()
    chrome_browser.auto_close=False
    url = "https://iss.fortaleza.ce.gov.br/grpfor/login.seam;jsessionid=1GC7+OZ2wq6sBwl0ZemSuXHB.pistol:iss-prod-02?cid=151891"
    chrome_browser.open_chrome_browser(url=url, preferences=chrome_prefs)
    while chrome_browser.does_page_contain_element('login:captchaDecor:captchaLogin'):
        chrome_browser.wait_until_element_is_visible('login:password')
        chrome_browser.input_text('login:username', '96287586320')
        chrome_browser.input_text('login:password', 'IEL000178@')
        chrome_browser.capture_element_screenshot('//*[@id="login:captchaDecor"]/img', 'captcha.png')
        chrome_browser.input_text('login:captchaDecor:captchaLogin',anti_captcha("captcha.png"))
        chrome_browser.click_element_if_visible('login:botaoEntrar')
    for nf in nfs:
        if type(nf) == int :
            chrome_browser.go_to('https://iss.fortaleza.ce.gov.br/grpfor/pages/nfse/consultaNfsePorNumero.seam?cid=212315')
            chrome_browser.input_text('consultarnfseForm:numNfse', nf)
            chrome_browser.click_element('consultarnfseForm:j_id227')
            chrome_browser.click_element_when_visible('//*[@id="consultarnfseForm:dataTable:0:j_id356"]/a')
            chrome_browser.click_element_when_visible('//*[@id="j_id153:panelAcoes"]/tbody/tr/td[1]/input')
            while not check_if_download_has_finished(default_download_directory):
                pass
            temp_filename_for_downloads(default_download_directory, nf)    
        else:
            pass    
    chrome_browser.close_browser()


def log_in_and_download_speedgov(download_dir: str, region: str, list_of_clients: list[ClientInfo], login: str, password:str):
    browser = Selenium()
    browser.auto_close=False
    home_dir = Path().home()
    download_dir = f"{home_dir}\\Documents\\Documentos RPA FIEC\\{region}"
    url_login = f"http://iss.speedgov.com.br/{region}/login"

    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }

    browser.open_chrome_browser(url_login, preferences=chrome_prefs)
    browser.input_text("inscricao", login)
    browser.input_text("xpath://*[@id='senha']", password)
    browser.click_element("xpath:/html/body/div[1]/div/div[1]/div[2]/form/button")
    for nf in list_of_clients:
        browser.go_to(f"http://iss.speedgov.com.br/{region}/pdfs/{nf.number}/notafiscal")
        check_if_download_has_finished(download_dir)
        time.sleep(0.5)
    time.sleep(2)
    browser.close_browser()





