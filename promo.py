import pyautogui
import time
import os
import openpyxl
import tkinter as tk
from tkinter import messagebox

def abrir_erp_login():
    # ABERTURA DO SISTEMA
    print("Abrindo ERP")
    os.startfile('C:\\Program Files\\Alterdata\\ERP\\Bimer.exe')
    time.sleep(5)

    # INSERIR USUARIO
    # pyautogui.click(950, 454)
    # pyautogui.write('MOISES')

    # INSERIR SENHA
    print("Inserindo Senha")
    pyautogui.click(945, 523)
    pyautogui.write('09042021')

    time.sleep(1)

    # ENTRAR
    print("Entrando")
    pyautogui.click(981, 631)
    time.sleep(1)

def extrair_relatorio():
    # EXPANDIR MENU
    pyautogui.click(163, 237)
    time.sleep(1)

    # CONSULTA ESTOQUE
    pyautogui.click(136,439)
    time.sleep(4)

    # INFORMAR CODIGO
    pyautogui.click(1526, 171)
    time.sleep(1)

def ler_proximo_codigo(sheet, coluna_origem, coluna_marcador):
    # Itera sobre as linhas da planilha
    for row in range(1, sheet.max_row + 1):
        # Verifica se a célula já foi lida
        if sheet[f'{coluna_marcador}{row}'].value != 'LIDO':
            # Lê a informação da célula de origem
            valor = sheet[f'{coluna_origem}{row}'].value

            # Marca a célula como lida
            sheet[f'{coluna_marcador}{row}'].value = 'LIDO'

            # Retorna o valor e a linha para marcar como lido depois
            return valor, row

    return None, None

def inserir_codigo_no_erp(valor):
    # Insere o valor no ERP usando pyautogui
    if valor:
        print(f"Inserindo valor: {valor}")
        pyautogui.write(str(valor))
        pyautogui.press('enter')  # Pesquisando item
        time.sleep(1)

def selecionando_item():
    # Selecionando item
    pyautogui.doubleClick(640, 245)
    time.sleep(3)

    # Complementar
    pyautogui.click(563, 398)
    time.sleep(1)

    # Caracteristica
    pyautogui.click(1117, 758)
    time.sleep(1)

    # Adicionar
    pyautogui.click(661, 363)
    time.sleep(1)

    # Informando caracteristica
    pyautogui.click(695, 517)
    pyautogui.write('12')
    pyautogui.press('tab')
    time.sleep(1)

    # Gravar caracteristica
    pyautogui.click(1091, 549)
    time.sleep(1)

    # Cancelar edição
    pyautogui.click(1236, 556)
    time.sleep(1)

    # Gravar Produto
    pyautogui.click(1294, 858)
    time.sleep(2)

def excluir_codigo_anterior():
    # Selecionar campo do código
    pyautogui.click(1526, 171)
    time.sleep(1)

    # Apagar o código anterior
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    time.sleep(1)

def fechar_sistema():
    # FECHANDO SISTEMA
    pyautogui.click(1902, 10)
    time.sleep(1)

    # CONFIRMANDO
    pyautogui.click(836, 554)

def main():
    abrir_erp_login()
    extrair_relatorio()
    time.sleep(2)  # Espera o relatório carregar

    # Carrega o arquivo Excel
    workbook = openpyxl.load_workbook('C:\\Users\\Administrador\\Desktop\\Loja\\Automacao_Promo\\CODIGOS.xlsx')
    sheet = workbook['Planilha1']

    # Define as colunas de origem e marcador
    coluna_origem = 'A' # Coluna que vai ser lida
    coluna_marcador = 'C'  # Coluna para marcar que a célula foi lida

    while True:
        valor, row = ler_proximo_codigo(sheet, coluna_origem, coluna_marcador)
        if valor is None:
            break  # Sai do loop se não houver mais códigos

        inserir_codigo_no_erp(valor)
        selecionando_item()

        # Salva o arquivo Excel após cada inserção
        workbook.save('C:\\Users\\Administrador\\Desktop\\Loja\\Automacao_Promo\\CODIGOS.xlsx')

        # Excluir o código anterior antes de inserir o próximo
        excluir_codigo_anterior()

    fechar_sistema()
    print("Finalizado")

if __name__ == "__main__":
    main()