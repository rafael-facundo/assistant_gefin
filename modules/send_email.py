from RPA.Email.ImapSmtp import ImapSmtp
from pathlib import Path
from RPA.FileSystem import FileSystem


def send_email(nf: str, boleto: str, recipient: str):
    gmail_account = "rfcosta@sfiec.org.br"
    gmail_app_password = "jiitxvjdnglplqxx"
    mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
    mail.authorize(account=gmail_account, password=gmail_app_password)
    mail.send_message(sender=gmail_account,
                        recipients= recipient,
                        subject="Email Test",
                        body="email test",
                        attachments=[nf, boleto])
        


def comparing_docs_and_send_email(nfs_directory: str, boletos_directory: str, email: str):
    file = FileSystem()
    nfs = file.list_files_in_directory(nfs_directory)
    boletos = file.list_files_in_directory(boletos_directory)

    for nf in nfs:
        nf_combination = str(nf)
        nf_combination = nf_combination.split("SESI DR\\")[1]
        nf_combination = nf_combination[-11:-4]
        for boleto in boletos:
            boleto = str(boleto)
            if nf_combination in boleto:
                print("ACHEI")
                print(f"{nf} {boleto}")
                send_email(nf, boleto, email)
