from RPA.Assistant import Assistant

def escrever(letras):
    print(letras)

def interface_activ():
    face = Assistant()
    face.add_drop_down("house_choice", ["BAIXAR NOTAS", "SEPARAR BOLETOS", "ENVIAR EMAIS"], label="ESCOLHA UMA CASA")
    face.add_submit_buttons(["EXECUTAR", "CANCELAR"])
    activiti = face.run_dialog(100, "teste")
    return activiti

def interface_options():
    face = Assistant()
    face.add_drop_down("options", ["SESI", "SENAI", "FEIC", "IEL"], label="ESCOLHA UMA CASA")
    face.run_dialog(100, "teste")

activiti = interface_activ()
if activiti.get("house_choice") == "BAIXAR NOTAS":
    interface_options()
