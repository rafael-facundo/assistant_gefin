from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
from pathlib import Path


def rename_pdf(directory: str):
    file = FileSystem()
    pdf = PDF()
    list_of_file_names = file.list_files_in_directory(directory)
    for file_adr in list_of_file_names:
        pdf.open_pdf(file_adr)
        rps_matches = pdf.find_text("subtext:do RPS")
        rps = rps_matches[0].neighbours[0]
        file_str = str(file_adr)
        nf_number = file_str[-8:-4]
        print(nf_number)
        print(rps)
        print('')
        print(file_str)
        print(file_str[:-16])
        print('')
        new_name = f"{file_str[:-16]}DR\\NF {nf_number} RPS {rps}.pdf"
        pdf.close_pdf()
        file.move_file(file_adr, new_name)
        