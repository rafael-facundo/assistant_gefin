import imports


directory = r'C:\\Users\\rfcosta\\Documents\\Documentos RPA FIEC\\maracanau'
list_of_client_info = imports.read_excel(r'C:\\Users\\rfcosta\\Desktop\\Palácio\\Códigos\\assistant_gefin\\Planilhas\\sesi_albano.xlsx')
imports.rename_pdf_sesi_albano(directory)