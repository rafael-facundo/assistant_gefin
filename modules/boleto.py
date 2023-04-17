from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from pathlib import Path

def get_number_rps(num_page, directory: str):
    pdf = PDF()
    pdf.open_pdf(directory) 
    number_nf = pdf.find_text("Nro.Documento", num_page, "down") 
    number_nf = number_nf[2].neighbours[0] 
    number_nf = number_nf.lstrip('0') 
    return number_nf 

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