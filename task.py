from RPA.Assistant import Assistant
import sys
from pathlib import Path
from RPA.Browser.Selenium import Selenium
from modules.boleto import open_pdf_and_save_pages
from modules.iss_website import log_in_and_download_sesi_dr
from modules.read_excel import excel_sesi_dr
from modules.rename import rename_pdf
from modules.send_email import send_email

# nova chave 2ab05e6bf0186f7724d613b158b7365f
# email_password jiitxvjdnglplqxx

def start_dialog(op_list: list[str]):
    dialog = Assistant()
    dialog.add_drop_down("website_choice",
                         op_list,
                         label="O que deseja fazer?")
    dialog.add_file_input("input_excel_file",
                          "Selecione uma planilha",
                          file_type="xlsx")
    dialog.add_file_input("add_pdf_file", 'Selecione um boleto', file_type="pdf")
    dialog.add_submit_buttons(["Executar", "Cancelar"])
    user_input = dialog.run_dialog(90000, title="Assistente - GEFIN", height=400, width=300)

    return user_input

op_list = [
    "SEPARAR BOLETOS", "SESI DR", "SESI Parangaba", 'SESI Barra', 'SESI NR Saúde (Centro)',
            'SENAI DR', 'SENAI AUA', 'SENAI AABMS', 'SENAI WDS', 'IEL', 'FIEC', 
            'SESI Albano', 'SESI Clube', 'SENAI CETAFR', 'SENAI ISTEMM', 'SESI Juazeiro', 
            'SENAI WCC', 'SESI Sobral', 'SENAI Sobral', "Enviar Notas e Boletos"
            ]


if __name__ == "__main__":
    user_input_dict = start_dialog(op_list)
    if user_input_dict.get("submit") != "Executar":
        sys.exit()
    else:

        # workbook_path = user_input_dict.get("input_excel_file")[0]
        # list_of_client_info = excel_sesi_dr(workbook_path)

        if user_input_dict.get("website_choice") == op_list[0]: #Separar boletos
            boleto_path = user_input_dict.get("add_pdf_file")[0]
            open_pdf_and_save_pages(boleto_path)

        elif user_input_dict.get("website_choice") == op_list[1]: #sesi dr
            workbook_path = user_input_dict.get("input_excel_file")[0]
            list_of_client_info = excel_sesi_dr(workbook_path)
            browser = Selenium
            region = "SESI DR"
            download_dir = f"{Path().home()}\\Documents\\NOTAS BAIXADAS"
            default_download_directory = f"{download_dir}\\{region}"
            list_of_nfs = []
            for nf in list_of_client_info:
                list_of_nfs.append(nf.number)
            log_in_and_download_sesi_dr(browser, default_download_directory, list_of_nfs)
            rename_pdf(default_download_directory)

        elif user_input_dict.get("website_choice") == op_list[19]: #enviar emails
            workbook_path = user_input_dict.get("input_excel_file")[0]
            list_of_client_info = excel_sesi_dr(workbook_path)
            list_of_emails = []
            for client in list_of_client_info:
                list_of_emails.append(client.email)
            for email in list_of_emails:
                send_email()