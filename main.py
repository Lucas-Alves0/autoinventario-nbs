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
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Função para parar o script via botão
def parar_script():
    os._exit(0)

# Thread para verificar ESC a qualquer momento
def verificar_esc():
    while True:
        if keyboard.is_pressed('esc'):
            parar_script()

# Inicia a thread que escuta o ESC
esc_thread = threading.Thread(target=verificar_esc, daemon=True)
esc_thread.start()

Janela = Tk()
Janela.title("Automação Inventário In House")
Janela.resizable(width=False, height=False)
Janela.geometry("300x160")
edge = LabelFrame(Janela, text="Opções", borderwidth=1, relief="solid")
edge.place(x=5, y=5, width=290, height=150)


def insere_contagem():
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
    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    # Entrar na tela de criação do inventário

    Janela.destroy()
    sleep(0.5)
    hotkey('alt', 'tab')

     # Abre Digitação de Estoque
    sleep(0.3)
    moveTo(x=852, y=357)
    click(x=852, y=357)

    # Abre Digitação de Itens Avulsos
    sleep(0.3)
    moveTo(x=785, y=356)
    click(x=785, y=356)

    for item, saldo in zip(df["ITEM"], df["SALDO"]):
        # Insere PN  no campo
        sleep(0.1)
        moveTo(x=585, y=430)
        click(x=585, y=430)
        write(str(item), interval=0.00001)
        hotkey('Tab')

        # Seleciona Fornecedor
        sleep(0.1)
        moveTo(x=928, y=431)
        click(x=928, y=431)
        write('In House', interval=0.00001)   
        hotkey('Enter')
        hotkey('Tab')

        # Cola Saldo
        sleep(0.1)
        write(str(saldo), interval=0.00001)
        hotkey('Enter')

        # Clica em Incliur item e Salva
        sleep(0.1)
        moveTo(x=954, y=805)
        click(x=954, y=805)
        hotkey('Enter')

    press('shift')
    alert("Processo Finalizado!")

def SAIR():
    Janela.destroy()
    os._exit(0)

alert('Lembre-se de atualizar o saldo dos relatórios antes de executar, caso não o tenha feito clique em "Sair" e atualize-o.')

Invisible_button = Label(Janela)
Invisible_button.grid(row=0, column=0, padx=100, pady=5)
Start_Button = Button(Janela, text="Inserir Contagem", width=23, bd=3, command=insere_contagem)
Start_Button.grid(row=1, column=0, padx=50, pady=7)
Exit_Button = Button(Janela, text="Sair", width=23, bd=3, command=SAIR)
Exit_Button.grid(row=2, column=0, padx=50, pady=7)

Janela.mainloop()