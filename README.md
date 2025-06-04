# AutoInventário no NBS

Este projeto é um bot automatizado para inserir dados no módulo de contagem do NBS, utilizando a biblioteca `pyautogui`. Ele lê dados de uma planilha Google Sheets e preenche automaticamente os campos necessários na tela do sistema.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `pyautogui`
  - `pymsgbox`
  - `pandas`
  - `tkinter`
  - `gspread`
  - `google-auth`
  - `keyboard`
  - `urllib3`

Instale as dependências com:

```sh
pip install -r requirements.txt
```

## Configuração

1. Altere o caminho do arquivo de credenciais.
2. Certifique-se de ter acesso à planilha Google Sheets com a aba "Lista de Contagem".

## Uso

1. Execute o script principal:

    ```sh
    python main.py
    ```

2. Utilize a interface gráfica para selecionar a ação desejada:
    - **Inserir Contagem**: Insere itens na tela de criação do inventário, conforme a planilha Google Sheets.
    - **Sair**: Fecha a aplicação.

3. Para interromper o script a qualquer momento, pressione a tecla `ESC`.

## Estrutura do Projeto

- [main.py](main.py): Script principal com a lógica do bot e interface gráfica.
- [requirements.txt](requirements.txt): Lista de dependências do projeto.

## Observações

- O script utiliza automação de mouse e teclado, portanto, não utilize o computador durante a execução.
- Certifique-se de atualizar os relatórios de saldos antes de executar o script.