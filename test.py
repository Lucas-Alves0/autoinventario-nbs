from pyautogui import *
from time import sleep


abre_digitacao = locateCenterOnScreen('Digitacao_contagem.png', confidence=0.8)
if abre_digitacao:
    moveTo(abre_digitacao)
    click(abre_digitacao)
else:
    print(f'Img not found on your screen.')

sleep(1)

abre_digitacao_avulsa = locateCenterOnScreen('Digitacao_de_item_avulso.png', confidence=0.8)
if abre_digitacao_avulsa:
    moveTo(abre_digitacao_avulsa)
    click(abre_digitacao_avulsa)
else:
    print(f'Img not found on your screen.')

sleep(1)

clica_inserir_pn = locateCenterOnScreen('Inserir_item.png', confidence=0.8)
if clica_inserir_pn:
    moveTo(clica_inserir_pn)
    click(clica_inserir_pn)
else:
    print(f'Img not found on your screen.')

sleep(1)

seta_fornecedor = locateCenterOnScreen('Seta.png', confidence=0.8)
if seta_fornecedor:
    moveTo(seta_fornecedor)
    click(seta_fornecedor)
else:
    print(f'Img not found on your screen.')

sleep(1)

confirmar = locateCenterOnScreen('Confirmar_item.png', confidence=0.8)
if confirmar:
    moveTo(confirmar)
    click(confirmar)
else:
    print(f'Img not found on your screen.')