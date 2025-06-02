from pyautogui import *
from pymsgbox import alert
from time import sleep
from pandas import *
from tkinter import *
import keyboard
import os
import threading
import openpyxl


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
borda = LabelFrame(Janela, text="Opções", borderwidth=1, relief="solid")
borda.place(x=5, y=5, width=290, height=150)


def insere_contagem():
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


    picking = read_excel("Lista_de_Itens.xlsx", engine='openpyxl')
    for item, saldo in zip(picking["ITEM"], picking["SALDO"]):
        # Insere PN  no campo
        sleep(0.2)
        moveTo(x=585, y=430)
        click(x=585, y=430)
        write(str(item))
        hotkey('Tab')

        # Seleciona Fornecedor
        sleep(0.2)
        moveTo(x=928, y=431)
        click(x=928, y=431)
        write('In House')   
        hotkey('Enter')
        hotkey('Tab')

        # Cola Saldo
        sleep(0.2)
        write(str(saldo))
        hotkey('Enter')

        # Clica em Incliur item e Salva
        sleep(0.2)
        moveTo(x=954, y=805)
        click(x=954, y=805)
        hotkey('Enter')

    alert("Processo Finalizado!")

def SAIR():
    Janela.destroy()
    os._exit(0)


Edges = Label(Janela)
Edges.grid(row=0, column=0, padx=100, pady=5)
Button_1 = Button(Janela, text="Inserir Contagem", width=23, bd=3, command=insere_contagem)
Button_1.grid(row=1, column=0, padx=50, pady=7)
Button_2 = Button(Janela, text="Sair", width=23, bd=3, command=SAIR)
Button_2.grid(row=2, column=0, padx=50, pady=7)

Janela.mainloop()
