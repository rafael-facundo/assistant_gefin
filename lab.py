import imports


directory = r'C:\\Users\\rfcosta\\Documents\\Documentos RPA FIEC\\IEL'
list_of_client_info = imports.read_excel(r'C:\\Users\\rfcosta\\Desktop\\Palácio\\Códigos\\assistant_gefin\\Planilhas\\iel.xlsx')
imports.rename_pdf_iel(directory)