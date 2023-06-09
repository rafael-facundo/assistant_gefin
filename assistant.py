import imports
import sys
from pathlib import Path
from RPA.Browser.Selenium import Selenium

op_list = [
    "SEPARAR BOLETOS", "BAIXAR NOTAS: SESI DR", "BAIXAR NOTAS: SESI Parangaba", 'BAIXAR NOTAS: SESI Barra', 'BAIXAR NOTAS: SESI NR Saúde (Centro)',
            'BAIXAR NOTAS: SENAI DR', 'BAIXAR NOTAS: SENAI AUA', 'BAIXAR NOTAS: SENAI AABMS', 'BAIXAR NOTAS: SENAI WDS', 'BAIXAR NOTAS: IEL', 'BAIXAR NOTAS: FIEC', 
            'BAIXAR NOTAS: SESI Albano', 'BAIXAR NOTAS: SESI Clube', 'BAIXAR NOTAS: SENAI CETAFR', 'BAIXAR NOTAS: SENAI ISTEMM', 'BAIXAR NOTAS: SESI Juazeiro', 
            'BAIXAR NOTAS: SENAI WCC', 'BAIXAR NOTAS: SESI Sobral', 'BAIXAR NOTAS: SENAI Sobral', "ENVIAR NOTAS E BOLETOS"
            ]

user_input_dict = imports.start_dialog(op_list)
if user_input_dict.get("submit") != "Executar":
    sys.exit()
else:
    if user_input_dict.get("website_choice") == op_list[0]: #Separar boletos
        boleto_path = user_input_dict.get("add_pdf_file")[0]
        imports.open_pdf_and_save_pages(boleto_path)
        imports.alert("Todos os boletos foram separados.", 200, 430)

    elif user_input_dict.get("website_choice") == op_list[1]: #sesi dr
        workbook_path = user_input_dict.get("input_excel_file")[0]
        list_of_client_info = imports.read_excel(workbook_path)
        browser = Selenium
        region = "SESI DR"
        download_dir = f"{Path().home()}\\Documents\\Documentos RPA FIEC"
        default_download_directory = f"{download_dir}\\{region}"
        list_of_nfs = []
        for nf in list_of_client_info:
            list_of_nfs.append(nf.number)
        imports.log_in_and_download_sesi_dr(browser, default_download_directory, list_of_nfs)
        imports.rename_pdf_sesi_dr(default_download_directory)
        imports.alert("Todos as notas foram baixadas.", 200, 430)

    elif user_input_dict.get("website_choice") == op_list[9]: # iel
        workbook_path = user_input_dict.get("input_excel_file")[0]
        list_of_client_info = imports.read_excel(workbook_path)
        browser = Selenium
        region = "IEL"
        download_dir = f"{Path().home()}\\Documents\\Documentos RPA FIEC"
        default_download_directory = f"{download_dir}\\{region}"
        list_of_nfs = []
        for nf in list_of_client_info:
            list_of_nfs.append(nf.number)
        imports.log_in_and_download_iel(browser, default_download_directory, list_of_nfs)
        imports.rename_pdf_iel(default_download_directory)
        imports.alert("Todos as notas foram baixadas.", 200, 430)

    elif user_input_dict.get("website_choice") == op_list[11]: #sesi albano
        workbook_path = user_input_dict.get("input_excel_file")[0]
        list_of_client_info = imports.read_excel(workbook_path)
        browser = Selenium
        browser.auto_close=False
        region = "maracanau"
        download_dir = f"{Path().home()}\\Documents\\\\Documentos RPA FIEC"
        default_download_directory = f"{download_dir}\\{region}"
        imports.log_in_and_download_speedgov(default_download_directory, region, list_of_client_info, "1241313", "380901")
        imports.rename_pdf_sesi_albano(default_download_directory)
        imports.alert("Todos as notas foram baixadas.", 200, 430)

    elif user_input_dict.get("website_choice") == op_list[19]: #enviar emails
        boleto_directory = f"{Path().home()}\\Documents\\Documentos RPA FIEC\\boletos separados"
        nfs_directory = user_input_dict.get("add_pdf_directory")[0]
        nfs_directory = str(nfs_directory).split("\\NF")[0]
        workbook_path = user_input_dict.get("input_excel_file")[0]
        list_of_client_info = imports.read_excel(workbook_path)
        imports.send_email(list_of_client_info, nfs_directory, boleto_directory)
        imports.alert("Todos os envios foram concluídos.", 200, 430)
    






    

