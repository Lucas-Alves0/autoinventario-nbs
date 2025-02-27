# Automatizador de Inventário no NBS 

Este projeto é um bot automatizado para inserir dados no módulo de contagem do nbs, utilizando a biblioteca `pyautogui`. Ele lê dados de um arquivo Excel e preenche automaticamente os campos necessários na tela.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `pyautogui`
  - `pymsgbox`
  - `pandas`
  - `tkinter`
  - `openpyxl`

## Uso

1. Coloque o arquivo `Lista_de_Itens.xlsx` no diretório [Downloads](http://_vscodecontentref_/1).
2. Execute o script [automatizador-inventario-nbs.py](http://_vscodecontentref_/2):
    ```sh
    python automatizador-inventario-nbs.py
    ```
3. Utilize a interface gráfica para selecionar a ação desejada:
    - **Inserir Itens (Contagem)**: Insere itens na tela de criação do inventário.
    - **Inserir Dados (Contagem)**: Insere dados na tela de inserção de itens.
    - **Sair**: Fecha a aplicação.

## Estrutura do Projeto

- [automatizador-inventario-nbs.py](http://_vscodecontentref_/3): Script principal que contém a lógica do bot e a interface gráfica.
- `Lista_de_Itens.xlsx`: Arquivo Excel contendo os itens a serem inseridos.