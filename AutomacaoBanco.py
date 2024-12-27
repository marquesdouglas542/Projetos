import pyautogui
import time
import os
import sqlite3
from cryptography.fernet import Fernet
from PIL import Image, ImageEnhance, ImageFilter
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

pyautogui.PAUSE = 0.3

relatorio = 'PSG_Relatorio com itens - (0097)'
salvar_relatorio_no_diretorio = '\\\\tsclient\\C\\Users\\JFC\\Desktop\\Relatorios_Psg'
# Definindo o diretório das imagens
diretorio_imagens = os.path.join(os.getcwd(), 'posicoesPorImagem')

# Função para obter o caminho completo da imagem
def obter_caminho_imagem(nome_imagem):
    return os.path.join(diretorio_imagens, nome_imagem)

# Função para carregar e melhorar a imagem com Pillow
def melhorar_imagem(imagem_path):
    # Carregar a imagem usando Pillow
    img = Image.open(imagem_path)
    
    # Aumentar o contraste
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2)  # Aumenta o contraste para 2x

    # Aumentar o brilho
    enhancer = ImageEnhance.Brightness(img)
    img = enhancer.enhance(1.2)  # Aumenta o brilho em 20%

    # Aplicar um filtro de nitidez
    img = img.filter(ImageFilter.SHARPEN)

    # Salvar a imagem otimizada como uma versão temporária
    temp_path = imagem_path.replace(".png", "_melhorada.png")
    img.save(temp_path)
    
    return temp_path

# Função de localização com melhora de imagem
def locate_with_enhanced_image(imagem_name, confidence=0.9, retries=3, timeout=5):
    # Obter o caminho completo da imagem
    imagem_path = obter_caminho_imagem(imagem_name)
    
    # Melhorar a imagem usando Pillow
    imagem_melhorada_path = melhorar_imagem(imagem_path)
    
    # Tentativas de localização da imagem
    for attempt in range(retries):
        try:
            # Tentar localizar a imagem otimizada
            localização = pyautogui.locateOnScreen(imagem_melhorada_path, confidence=confidence)
            
            if localização:
                return localização
            else:
                print(f"Tentativa {attempt + 1} falhou para a imagem {imagem_name}. Tentando novamente...")
                time.sleep(timeout)
        except pyautogui.ImageNotFoundException:
            print(f"Imagem {imagem_name} não encontrada. Tentando novamente...")
            time.sleep(timeout)
    
    # Se a imagem não for encontrada após as tentativas, levantar uma exceção
    raise Exception(f"Falha ao localizar a imagem {imagem_name} após {retries} tentativas.")

# Função para exibir uma janela de alerta
def exibir_alerta():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showinfo(
        "Aviso",
        "Por favor, não mexa no mouse ou no teclado durante a execução do processo.\n"
        "Certifique-se de que o Caps Lock esteja DESATIVADO.\n\n"
        "Clique em OK para continuar."
    )
    root.destroy()  # Fecha a janela após o alerta

# Função para abrir o SAP
def abrir_sap():
    pyautogui.hotkey('winleft', 'd')  # Garantir que vai estar na área de trabalho
    time.sleep(1)
    try:
        pyautogui.hotkey('winleft')
        pyautogui.write('Documentos: SAP - WINDOWS 10.rdp')
        time.sleep(2)
        pyautogui.hotkey('enter')
    except Exception as e:
        print(f"Erro ao abrir o SAP: {e}")
        salvar_log(f"Erro ao abrir o SAP: {e}")

# Função para logar no SAP
def logar_sap(s_Amb):
    time.sleep(1)
    try:
        ambiente = locate_with_enhanced_image('s_Amb.png', confidence=0.9)
        pyautogui.moveTo(ambiente)
        pyautogui.click()
        pyautogui.write(s_Amb)
        pyautogui.hotkey('enter')
    except Exception as e:
        print(f"Erro ao logar no SAP: {e}")
        salvar_log(f"Erro ao logar no SAP: {e}")

# Função para logar com usuário e senha
def logar_usuario(u, s):
    time.sleep(15)
    try:
        logar = locate_with_enhanced_image('loginSAP.png', confidence=0.9)
        pyautogui.moveTo(logar)
        pyautogui.click()
        pyautogui.write(u)
        
        time.sleep(2)
        
        pyautogui.hotkey('tab')
        pyautogui.write(s)
        pyautogui.hotkey('enter')
        
        time.sleep(10)
        
        Bcancelar = locate_with_enhanced_image('botaoCancelar.png', confidence=0.9)
        pyautogui.moveTo(Bcancelar)
        
        time.sleep(3)
        
        pyautogui.click(button='left')
        
        pyautogui.hotkey('esc')
    except Exception as e:
        print(f"Erro ao logar o usuário: {e}")
        salvar_log(f"Erro ao logar o usuário: {e}")

# Função para abrir a consulta de formulários
def abrir_consulta_formularios():
    time.sleep(3)
    try:
        icon_formulario = locate_with_enhanced_image('iconeFormularios.png', confidence=0.85)
        pyautogui.moveTo(icon_formulario)
        pyautogui.click()
        time.sleep(3)
        
        campo_formulario = locate_with_enhanced_image('campoFormulario.png', confidence=0.9)
        pyautogui.click(campo_formulario)
        pyautogui.write(relatorio)
        pyautogui.hotkey("enter")
        
        pyautogui.moveTo(x=1011, y=334)  # Início da barra de rolagem
        pyautogui.mouseDown(button="left", x=1010, y=478)
        time.sleep(1)
        pyautogui.moveTo(x=430, y=466)
        
        nome_relatorio = locate_with_enhanced_image('PSG_Relatorio com itens - (0097).png', confidence=0.95)
        
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
    except Exception as e:
        print(f"Erro ao abrir o formulário: {e}")
        salvar_log(f"Erro ao abrir o formulário: {e}")

# Função para salvar o relatório como XLSX
def salvar_relatorio_xlsx():
    try:
        botao_converter = locate_with_enhanced_image('converterXLSX.png', confidence=0.95)
        pyautogui.moveTo(botao_converter)
        pyautogui.click()
        
        time.sleep(2)
        
        mudar_diretorio = locate_with_enhanced_image('MudarDiretorio.png', confidence=0.8)
        pyautogui.moveTo(mudar_diretorio)
        pyautogui.click(button='right')
        botao_edit_address = locate_with_enhanced_image('botaoEditAddress.png', confidence=0.9)
        pyautogui.moveTo(botao_edit_address)
        pyautogui.click(button='left')
        pyautogui.write(salvar_relatorio_no_diretorio)
        pyautogui.hotkey('enter')
        time.sleep(2)
        
        selecionar_tipo_arquivo = locate_with_enhanced_image('barraDeSelecaoTipoArquivo.png', confidence=0.9)
        pyautogui.moveTo(selecionar_tipo_arquivo)
        pyautogui.click()
        time.sleep(1)
        
        tipo_xlsx = locate_with_enhanced_image('tipoXLSX.png', confidence=0.9)
        pyautogui.moveTo(tipo_xlsx)
        pyautogui.click()
        time.sleep(1)
        pyautogui.hotkey('enter')
        
        # Verificar se o arquivo já existe
        caminho_arquivo_destino = os.path.join(salvar_relatorio_no_diretorio, relatorio + '.xlsx')
        if os.path.exists(caminho_arquivo_destino):
            # Arquivo existe, então substitui
            os.remove(caminho_arquivo_destino)
            salvar_log(f"Arquivo {relatorio}.xlsx substituído em {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} devido a nova execução.")
        
        pyautogui.moveTo(x=423, y=378)
        time.sleep(2)
        pyautogui.click(button='left', clicks=2)
        time.sleep(1)
        pyautogui.press('esc')
        fechar_janela_final = locate_with_enhanced_image('fecharTelaFinal.png', confidence=0.7)
        pyautogui.moveTo(fechar_janela_final)
        pyautogui.click()
        fechar = locate_with_enhanced_image('fecharImagem.png', confidence=0.7)
        pyautogui.moveTo(fechar)
        pyautogui.click()
        pyautogui.hotkey('winleft', 'esc')
        pyautogui.PAUSE = 0.03
        pyautogui.press('esc')
        time.sleep(3)
        abrir_consulta_formularios()
        salvar_relatorio_xlsx()
    except Exception as e:
        print(f"Erro ao salvar o relatório: {e}")
        salvar_log(f"Erro ao salvar o relatório: {e}")

# Função para salvar o log no banco de dados SQLite
def salvar_log(mensagem):
    # Conectar ao banco de dados SQLite
    conn = sqlite3.connect('logs.db')
    cursor = conn.cursor()

    # Criar a tabela de logs caso não exista
    cursor.execute('''CREATE TABLE IF NOT EXISTS logs (
                        id INTEGER PRIMARY KEY,
                        mensagem TEXT,
                        data_hora TIMESTAMP)''')
    
    # Inserir a mensagem de log
    cursor.execute('INSERT INTO logs (mensagem, data_hora) VALUES (?, ?)', (mensagem, datetime.now()))
    conn.commit()
    conn.close()

# Função de descriptografar as credenciais
def descriptografar_credenciais():
    # Ler a chave de criptografia
    with open('chave.key', 'rb') as file:
        key = file.read()

    cipher_suite = Fernet(key)

    # Descriptografar as variáveis
    s_Amb = cipher_suite.decrypt(b'gAAAAABnabM3q1mQYK4m7OlADTqmHmIGlDsGKmajEdJ4PwKVjkx_rNK-mbYb5zqeHy9dKjJ-oGpRnxIeDM-4HYxdPwd-wHmlnlnvrPd0xiosU1_cKzWlRfo=').decode()
    u = cipher_suite.decrypt(b'gAAAAABnabM3SuHVPZSKOJPmdC92oOwS2PyZ8K7oIyOYI_vWzN2JGLpJE-_ltvqE_0VVZlvHCuK2UOZd4MdArqCU1J1OeHvpHw==').decode()
    s = cipher_suite.decrypt(b'gAAAAABnabM3vCVDGHMCLSA7GGnIPYaTbYXrdCdskePch4TWyAA0KRu3bHkPzQXa2XNXFix7Dx0bmmOCuDm_eCW23tZf9dpyDA==').decode()

    return s_Amb, u, s

# Função principal
if __name__ == "__main__":
    # Exibir alerta inicial
    exibir_alerta()

    # Descriptografar as credenciais
    s_Amb, u, s = descriptografar_credenciais()

    # Executar as funções principais
    abrir_sap()
    logar_sap(s_Amb)
    logar_usuario(u, s)
    abrir_consulta_formularios()
    salvar_relatorio_xlsx()
