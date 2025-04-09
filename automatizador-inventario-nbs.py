from pyautogui import *
from pymsgbox import alert
from time import sleep
from pandas import *
from tkinter import *
import openpyxl

Janela = Tk()
Janela.title("Bot De>Para NBS")
# Janela.iconbitmap('C:\\Users\\lucas.paula_kovi\\Downloads\\Icone.ico')
Janela.resizable(width=False, height=False)
Janela.geometry("300x220")
borda = LabelFrame(Janela, text="Opções", borderwidth=1, relief="solid")
borda.place(x=5, y=5, width=290, height=210)


def contagem_avulsa():
    # Entrar na tela de criação do inventário

    Janela.destroy()
    sleep(1.5)
    hotkey('alt', 'tab')

    # Inserir Itens

    picking = read_excel("Lista_de_Itens.xlsx", engine='openpyxl')
    for item in (picking["ITEM"]):
        # Abrir Inserção de Itens

        sleep(1)
        moveTo(x=760, y=359)
        sleep(0.2)
        click(x=760, y=359)

        # Inserir itens

        sleep(0.2)
        moveTo(x=912, y=493)
        sleep(0.2)
        click(x=912, y=493)
        sleep(0.2)
        write('In House')
        sleep(0.2)
        hotkey('Enter')
        hotkey('tab')
        write('INHOUSE')
        sleep(0.2)
        moveTo(x=732, y=597)
        sleep(0.2)
        click(x=732, y=597)
        sleep(0.2)
        moveTo(x=1062, y=619)
        sleep(0.2)
        click(x=1062, y=619)
        sleep(0.2)
        write('in')
        sleep(0.2)
        hotkey('Enter')
        sleep(0.2)
        moveTo(x=749, y=433)
        sleep(0.2)
        click(x=749, y=433)
        sleep(0.2)
        moveTo(x=798, y=736)
        sleep(0.2)
        click(x=798, y=736)
        sleep(0.2)
        write(item)
        sleep(0.2)
        hotkey('tab')
        sleep(0.2)
        moveTo(x=1143, y=739)
        sleep(0.2)
        click(x=1143, y=739)
        sleep(0.2)
        write('In House')
        sleep(0.2)
        hotkey('Enter')
        sleep(0.2)
        hotkey('Enter')
        sleep(0.2)
        hotkey('Enter')
        sleep(0.2)
        hotkey('Enter')
        sleep(0.2)

    alert("Processo Finalizado!")

def Dados_Contagem():
    # Entrar na tela de inserção de itens

    Janela.destroy()
    sleep(1.5)
    hotkey('alt', 'tab')

    # Abrir tela de inserção

    sleep(1)
    moveTo(x=851, y=358)
    click(x=851, y=358)
    sleep(0.1)
    moveTo(x=525, y=800)
    click(x=525, y=800)
    sleep(0.1)
    moveTo(x=1107, y=405)
    click(x=1107, y=405)
    sleep(0.1)
    moveTo(x=1217, y=527)
    click(x=1217, y=527)

    picking = read_excel("Lista_de_Itens.xlsx", engine='openpyxl')
    for saldo in (picking["SALDO"]):
        write(str(saldo))
        hotkey('Enter')

    alert('Finalizado, favor verifique!')

def SAIR():
    Janela.destroy()


B = Label(Janela, text="")
B.grid(row=0, column=0, padx=100, pady=5)
B1 = Button(Janela, text="Inserir Itens (Contagem)", width=23, bd=3, command=contagem_avulsa)
B1.grid(row=1, column=0, padx=50, pady=7)
B2 = Button(Janela, text="Inserir Dados (Contagem)", width=23, bd=3, command=Dados_Contagem)
B2.grid(row=2, column=0, padx=50, pady=7)
B3 = Button(Janela, text="Sair", width=23, bd=3, command=SAIR)
B3.grid(row=3, column=0, padx=50, pady=7)

Janela.mainloop()
