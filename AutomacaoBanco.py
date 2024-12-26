import pyautogui
import time
import os
from cryptography.fernet import Fernet

# Definindo o diretório das imagens
diretorio_imagens = os.path.join(os.getcwd(), 'posicoesPorImagem')

# Função para obter o caminho completo da imagem
def obter_caminho_imagem(nome_imagem):
    return os.path.join(diretorio_imagens, nome_imagem)

relatorio = 'PSG_Relatorio com itens - (0097)'
salvar_relatorio_no_diretorio = '\\\\tsclient\\C\\Users\\JFC\\Desktop\\Relatorios_Psg'

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

def abrir_sap():
    pyautogui.hotkey('winleft', 'd')  # Garantir que vai estar na área de trabalho
    time.sleep(1)   
    iconeSAP = pyautogui.locateOnScreen(obter_caminho_imagem('iconeSAP.png'), confidence=0.8)
    pyautogui.moveTo(iconeSAP)
    pyautogui.click()
    pyautogui.hotkey('enter')

def logar_sap(s_Amb):
    time.sleep(1)
    ambiente = pyautogui.locateOnScreen(obter_caminho_imagem('s_Amb.png'), confidence=0.8)
    pyautogui.moveTo(ambiente)
    pyautogui.click()
    pyautogui.write(s_Amb)
    pyautogui.hotkey('enter')
    
def logar_usuario(u, s):
    time.sleep(15)
    logar = pyautogui.locateOnScreen(obter_caminho_imagem('loginSAP.png'), confidence=0.8)
    pyautogui.moveTo(logar)
    pyautogui.click()
    pyautogui.write(u)
    
    time.sleep(2)
    
    pyautogui.hotkey('tab')
    pyautogui.write(s)
    pyautogui.hotkey('enter')
    
    time.sleep(10)
    
    Bcancelar = pyautogui.locateOnScreen(obter_caminho_imagem('botaoCancelar.png'), confidence=0.8)
    pyautogui.moveTo(Bcancelar)
    
    time.sleep(3)
    
    pyautogui.click(button='left')
    
    pyautogui.hotkey('esc')
  
def abrir_consulta_formularios():
    time.sleep(3)
    icon_formulario = pyautogui.locateOnScreen(obter_caminho_imagem('iconeFormularios.png'), confidence=0.9)
    pyautogui.moveTo(icon_formulario)
    pyautogui.click()
    time.sleep(3)
    
    campo_formulario = pyautogui.locateOnScreen(obter_caminho_imagem('campoFormulario.png'), confidence=0.9)
    pyautogui.click(campo_formulario)
    pyautogui.write(relatorio)
    pyautogui.hotkey("enter") 
    
    pyautogui.moveTo(x=1011, y=334)  # Início da barra de rolagem
    pyautogui.mouseDown(button="left", x=1010, y=478)
    time.sleep(1)
    pyautogui.moveTo(x=430, y=466)
    
    nome_relatorio = pyautogui.locateOnScreen(obter_caminho_imagem('PSG_Relatorio com itens - (0097).png'), confidence=0.85)
    
    if nome_relatorio == relatorio.strip('.png'): 
        pyautogui.moveTo(nome_relatorio)
        pyautogui.click(button='left')
        pyautogui.press("enter")
        time.sleep(3)
    else: 
        pyautogui.moveTo(nome_relatorio)
        pyautogui.click(button='left')
        pyautogui.press("enter")
        time.sleep(3)
        
def salvar_relatorio_xlsx():
    botao_converter = pyautogui.locateOnScreen(obter_caminho_imagem('converterXLSX.png'), confidence=0.9)
    pyautogui.moveTo(botao_converter)
    pyautogui.click()
    
    time.sleep(2)
    
    mudar_diretorio = pyautogui.locateOnScreen(obter_caminho_imagem('MudarDiretorio.png'), confidence=0.9)
    pyautogui.moveTo(mudar_diretorio)
    pyautogui.click(button='right')
    botao_edit_address = pyautogui.locateOnScreen(obter_caminho_imagem('botaoEditAddress.png'), confidence=0.9)
    pyautogui.moveTo(botao_edit_address)
    pyautogui.click(button='left')
    pyautogui.write(salvar_relatorio_no_diretorio)
    pyautogui.hotkey('enter')
    time.sleep(2)
    
    selecionar_tipo_arquivo = pyautogui.locateOnScreen(obter_caminho_imagem('barraDeSelecaoTipoArquivo.png'), confidence=0.9)
    pyautogui.moveTo(selecionar_tipo_arquivo)
    pyautogui.click()
    time.sleep(1)
    
    tipo_xlsx = pyautogui.locateOnScreen(obter_caminho_imagem('tipoXLSX.png'), confidence=0.9)
    pyautogui.moveTo(tipo_xlsx)
    pyautogui.click()
    time.sleep(1)
    pyautogui.hotkey('enter')
    
    pyautogui.moveTo(x=422, y=391)
    time.sleep(2)
    pyautogui.click(button='left', clicks=2)
    time.sleep(1)
    pyautogui.press('esc')
    fechar_janela_final = pyautogui.locateOnScreen(obter_caminho_imagem('fecharTelaFinal.png'), confidence=0.9)
    pyautogui.moveTo(fechar_janela_final)
    pyautogui.click()
    pyautogui.press('esc')
    time.sleep(3)
    abrir_consulta_formularios()
    salvar_relatorio_xlsx()

if __name__ == "__main__":
    abrir_sap()
    logar_sap(s_Amb)
    logar_usuario(u, s)
    abrir_consulta_formularios()
    salvar_relatorio_xlsx()
