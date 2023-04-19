from RPA.Assistant import Assistant
import sys
from pathlib import Path
from RPA.Browser.Selenium import Selenium
from modules.boleto import open_pdf_and_save_pages
from modules.iss_website import log_in_and_download_sesi_dr
from modules.read_excel import excel_sesi_dr
from modules.rename import rename_pdf
# from modules.send_email import send_email_2
from modules. read_excel import ClientInfo
from RPA.Email.ImapSmtp import ImapSmtp

# nova chave 2ab05e6bf0186f7724d613b158b7365f
# email_password jiitxvjdnglplqxx


def send_email_2(list_of_client_info: list[ClientInfo], nfs_directory: str, boletos_directory: str):
    for info in list_of_client_info:
        nf = f"{nfs_directory}\\NF {info.number} RPS {info.rps}.pdf"
        boleto = f"{boletos_directory}\\RPS {info.rps} BOLETO.pdf"
        gmail_account = "rfcosta@sfiec.org.br"
        gmail_app_password = "jiitxvjdnglplqxx"
        mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
        mail.authorize(account=gmail_account, password=gmail_app_password)
        mail.send_message(sender=gmail_account,
                            recipients= info.email,
                            subject="Email Test",
                            body="email test",
                            attachments=[nf, boleto])


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
            # count = 0
            boleto_directory = f"{Path().home()}\\Documents\\Documentos RPA FIEC\\boletos separados"
            nfs_directory = user_input_dict.get("add_pdf_directory")[0]
            nfs_directory = str(nfs_directory).split("\\NF")[0]
            workbook_path = user_input_dict.get("input_excel_file")[0]
            list_of_client_info = excel_sesi_dr(workbook_path)
            send_email_2(list_of_client_info, nfs_directory, boleto_directory)
            #comparing_docs_and_send_email(nfs_directory, boleto_directory, list_of_client_info)
            # list_of_emails = []
            # for client in list_of_client_info:
            #     list_of_emails.append(client.email)
            # for email in list_of_emails:
            #     if type(email) == type("None"):
            #         # print(email)
            #         # print(type(email))
            #         comparing_docs_and_send_email(nfs_directory, boleto_directory, email)
            #         count += 1
            #     if count == 1:
            #         break 