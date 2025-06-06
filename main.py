from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials
from pymsgbox import alert
from pyautogui import *
from time import sleep
from tkinter import *
import pandas as pd
import threading
import keyboard
import warnings
import gspread
import urllib3
import ctypes
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Verifica se o Caps está ativo, e o desliga
def caps_lock_check():
    return ctypes.WinDLL('User32.dll').GetKeyState(0x14) & 0x0001

def verifica_caps():
    while True:
        if caps_lock_check():
            press('capslock')

# Função para parar o script via botão
def parar_script():
    os._exit(0)

# Thread para verificar ESC a qualquer momento
def verificar_esc():
    while True:
        if keyboard.is_pressed('esc'):
            parar_script()

# Inicia a thread que executa o ESC
esc_thread = threading.Thread(target=verificar_esc, daemon=True)
esc_thread.start()

# Inicia a thread que desativa o Caps
caps_thread = threading.Thread(target=verifica_caps, daemon=True)
caps_thread.start()

Janela = Tk()
Janela.title("Automação Inventário In House")
Janela.resizable(width=False, height=False)
Janela.geometry("300x160")
edge = LabelFrame(Janela, text="Opções", borderwidth=1, relief="solid")
edge.place(x=5, y=5, width=290, height=150)
secs_b_keys = 0.00000001  # Intervalo entre teclas

# Abre Conecta ao Google Sheets
scope = [
"https://www.googleapis.com/auth/spreadsheets",
"https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(r"C:\Users\lucas.paula_kovi\VSCodeProjects\creds.json", scopes=scope)
session = AuthorizedSession(creds)
session.verify = False
client = gspread.authorize(creds, session=session)

sheet = client.open_by_key("1HzZdqDwhg0YJcvOrQA9Zu9iH1MZGbGheJIvtR76Lupo").worksheet("Lista de Contagem")
cont_pns = len(sheet.col_values(1)[1:])
data = sheet.get_all_records()
df = pd.DataFrame(data)

def insere_contagem():
    # Entrar na tela de criação do inventário
    try:
        Janela.destroy()
        sleep(0.5)
        hotkey('alt', 'tab')
        sleep(1)

        # Abre Digitação de Estoque
        abre_digitacao = locateCenterOnScreen(r'buttons\Digitacao_contagem.png', confidence=0.8)
        moveTo(abre_digitacao)
        click(abre_digitacao)
        sleep(1)

        # Abre Digitação de Itens Avulsos
        abre_digitacao_avulsa = locateCenterOnScreen(r'buttons\Digitacao_de_item_avulso.png', confidence=0.8)
        moveTo(abre_digitacao_avulsa)
        click(abre_digitacao_avulsa)
        sleep(1)

        for item, saldo in zip(df["ITEM"], df["SALDO"]):
            # Insere PN  no campo
            sleep(0.3)
            insere_item = locateCenterOnScreen(r'buttons\Inserir_item.png', confidence=0.8)
            moveTo(insere_item)
            click(insere_item)
            write(str(item), interval=secs_b_keys)
            hotkey('Tab')

            # Seleciona Fornecedor
            sleep(0.3)
            seta_fornecedor = locateCenterOnScreen(r'buttons\Seta.png', confidence=0.8)
            moveTo(seta_fornecedor)
            click(seta_fornecedor)
            write('In H', interval=secs_b_keys)   
            hotkey('Enter')
            hotkey('Tab')

            # Cola Saldo
            sleep(0.1)
            write(str(saldo), interval=secs_b_keys)
            hotkey('Enter')

            # Clica em Incliur item e Salva
            sleep(0.3)
            confirmar = locateCenterOnScreen(r'buttons\Confirmar_item.png', confidence=0.8)
            moveTo(confirmar)
            click(confirmar)
            hotkey('Enter')

        press('shift')
        alert("Processo Finalizado!")
    except:
        alert('Botões do NBS não encontrados na tela')
        os._exit(0)

def SAIR():
    Janela.destroy()
    os._exit(0)

alert('Lembre-se de atualizar o saldo dos relatórios antes de executar, caso não o tenha feito clique em "Sair" e atualize-o.')
alert(f'A lista contém {cont_pns} PN(s) e o tempo estimado para conclusão é de  {(cont_pns * 10) // 60}min(s){(cont_pns * 10) % 60}s.')

Invisible_button = Label(Janela)
Invisible_button.grid(row=0, column=0, padx=100, pady=5)
Start_Button = Button(Janela, text="Inserir Contagem", width=23, bd=3, command=insere_contagem)
Start_Button.grid(row=1, column=0, padx=50, pady=7)
Exit_Button = Button(Janela, text="Sair", width=23, bd=3, command=SAIR)
Exit_Button.grid(row=2, column=0, padx=50, pady=7)

Janela.mainloop()