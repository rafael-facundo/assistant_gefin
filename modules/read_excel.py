from RPA.Excel.Files import Files

class ClientInfo:
    name: str
    number: int
    rps: int
    body: str
    email: str


    def __init__(self, name: str, number: int, rps: int, email: str) -> None:
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

def excel_sesi_dr(excel_file_path: str):
    excel = Files()
    excel.open_workbook(excel_file_path)
    excel_dict = excel.read_worksheet(header=True)
    list_of_excel = []
    for line in excel_dict:
        list_of_excel.append(ClientInfo(line.get("RAZÃO SOCIAL"), line.get("NOTA FISCAL"), line.get("RPS (REC PROV. SERVIÇO)GER"), line.get("E-MAILS DE CONTATO")))
    return list_of_excel