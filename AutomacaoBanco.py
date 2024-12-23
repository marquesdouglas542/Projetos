import pyautogui
import time
from cryptography.fernet import Fernet

# Tempo de aguardo entre cada ação
pyautogui.PAUSE = 0.3

# Ler a chave de criptografia do arquivo
with open('chave.key', 'rb') as file:
    key = file.read()

# Criar objeto Fernet com a chave
cipher_suite = Fernet(key)

# Senhas criptografadas 
HYxdPwd = b'gAAAAABnabM3q1mQYK4m7OlADTqmHmIGlDsGKmajEdJ4PwKVjkx_rNK-mbYb5zqeHy9dKjJ-oGpRnxIeDM-4HYxdPwd-wHmlnlnvrPd0xiosU1_cKzWlRfo='
JPmdC = b'gAAAAABnabM3SuHVPZSKOJPmdC92oOwS2PyZ8K7oIyOYI_vWzN2JGLpJE-_ltvqE_0VVZlvHCuK2UOZd4MdArqCU1J1OeHvpHw=='
TWyAA = b'gAAAAABnabM3vCVDGHMCLSA7GGnIPYaTbYXrdCdskePch4TWyAA0KRu3bHkPzQXa2XNXFix7Dx0bmmOCuDm_eCW23tZf9dpyDA=='

# D.S
s_Amb = cipher_suite.decrypt(HYxdPwd).decode()
u = cipher_suite.decrypt(JPmdC).decode()
s = cipher_suite.decrypt(TWyAA).decode()

# Definindo o diretório e relatório
relatorio = "PSG_Relatorio com itens - (0097)"
diretorio = "\\\\tsclient\\C\\Users\\JFC\\Desktop\\Relatorios_Psg"  # Mudar conforme necessário

def abrir_sap():
    pyautogui.click(x=1362, y=751)  # Garantir que vai estar na área de trabalho
    pyautogui.moveTo(x=555, y=25, duration=1)  # Pegar posição do SAP na desktop (Alterar conforme necessário)
    pyautogui.click(x=555, y=25)  # Clicar no ícone (alterar conforme necessário)
    pyautogui.press("enter")  # Para abrir o SAP
    time.sleep(5)

def logar_sap(s_Amb):
    pyautogui.moveTo(x=490, y=225, duration=1)  # Move o mouse até o campo de s
    pyautogui.click(x=490, y=225)  # Entra no campo de s
    pyautogui.write(s_Amb)  # Digita a s do ambiente (alterar conforme necessário)
    pyautogui.press("enter")  # Loga no ambiente
    time.sleep(15)  # Aguarda 15 segundos para logar no usuário, garantindo que o ambiente conectou

def logar_usuario(u, s):
    pyautogui.moveTo(x=793, y=364, duration=1)
    pyautogui.click(x=793, y=364)
    pyautogui.write(u)  # Usuário, mudar conforme necessário
    pyautogui.press("tab")  # Muda para o campo de s
    pyautogui.write(s)  # s do usuário, trocar conforme necessário
    pyautogui.press("enter")
    time.sleep(10)

def fechar_janela_selecionar_filial():
    pyautogui.moveTo(x=515, y=425, duration=1)
    pyautogui.click(x=515, y=425)
    pyautogui.press("esc")  # Fecha outra possível janela que abre
    time.sleep(5)

def abrir_gerenciador_consultas(relatorio):
    pyautogui.moveTo(x=860, y=53, duration=1)
    pyautogui.click(x=860, y=53)
    pyautogui.moveTo(x=632, y=264, duration=1)  # Campo de pesquisa
    pyautogui.click(x=632, y=264)
    pyautogui.write(relatorio)  # Nome da consulta (Alterar conforme necessário)
    pyautogui.moveTo(x=1011, y=334, duration=1)  # Início da barra de rolagem
    pyautogui.mouseDown(button="left", x=1010, y=478)
    time.sleep(5)
    pyautogui.moveTo(x=430, y=466, duration=1)
    pyautogui.click(x=430, y=466)
    pyautogui.press("enter")
    time.sleep(5)

def converter_xlsx(diretorio):
    pyautogui.moveTo(x=166, y=55, duration=1)
    pyautogui.click(x=166, y=55)
    time.sleep(5)
    pyautogui.moveTo(x=893, y=161, duration=1)
    pyautogui.click(x=893, y=161)
    pyautogui.write(diretorio)
    time.sleep(5)
    pyautogui.press("enter")
    time.sleep(5)
    pyautogui.moveTo(x=614, y=509, duration=1)
    pyautogui.click(x=614, y=509)
    pyautogui.moveTo(x=565, y=564, duration=1)
    pyautogui.click(x=565, y=564)
    time.sleep(5)
    pyautogui.press("enter")
    pyautogui.moveTo(x=422, y=381, duration=1)
    pyautogui.click(x=422, y=381)
    time.sleep(3)
    pyautogui.press("enter")
    time.sleep(6)
    pyautogui.click(x=681, y=317)
    pyautogui.press("esc")
    
def fechar_relatorio():
    time.sleep(3)
    pyautogui.click(x=681, y=317)
    pyautogui.press("esc")

if __name__ == "__main__":
    abrir_sap()
    logar_sap(s_Amb)
    logar_usuario(u, s)
    fechar_janela_selecionar_filial()  # Dependendo do usuário, esse trecho pode ser comentado
    abrir_gerenciador_consultas(relatorio)
    converter_xlsx(diretorio)
    fechar_relatorio()
