import openpyxl

#colocando os dados com acénto e parte maiúscula
import pyperclip
#colocando uma "pausa" para carregar a página
from time import sleep

import pyautogui

workSpace = openpyxl.load_workbook('produtos_ficticios.xlsx')
workSheet = workSpace['Produtos']
for linha in workSheet.iter_rows(min_row=2 , max_row=2):
   
    nomeProduto = linha[0].value
    pyperclip.copy(nomeProduto)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1472,413,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1467,526,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigoProduto = linha[3].value
    pyperclip.copy(codigoProduto)
    pyautogui.click(1481,614,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1463,691,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1465,783,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Trocar a página
    pyautogui.click(1456,854,duration=1)
    #pausa
    sleep(5)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1539,325,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidadeEstoque = linha[7].value
    pyperclip.copy(quantidadeEstoque)
    pyautogui.click(1515,411,duration=1)
    pyautogui.hotkey('ctrl','v')

    dataValidade = linha[8].value
    pyperclip.copy(dataValidade)
    pyautogui.click(1504,501,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1489,574,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Tamanho é diferente porque é um "select"/"Dropdown"
    tamanho = linha[10].value
    pyautogui.click(1530,670,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1520,705,duration=1)
    elif tamanho =='Médio':
        pyautogui.click(1509,730,duration=1)
    else:
        pyautogui.click(1503,748,duration=1)
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1496,754,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão próximo
    pyautogui.click(1494,808,duration=1)
    sleep(5)


    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1465,347,duration=1)
    pyautogui.hotkey('ctrl','v')

    paisOrigem = linha[13].value
    pyperclip.copy(paisOrigem)
    pyautogui.click(1473,426,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1475,523,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigoBarras = linha[15].value
    pyperclip.copy(codigoBarras)
    pyautogui.click(1471,655,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacaoArmazem = linha[16].value
    pyperclip.copy(localizacaoArmazem)
    pyautogui.click(1472,733,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão Concluir
    pyautogui.click(1485,814,duration=1)
    #Botão 1° confirmar
    pyautogui.click(1887,191,duration=1)
    #Botão 2° confirmar
    pyautogui.click(1886,191,duration=1)
    #Cadastrar novamente
    pyautogui.click(1695,580,duration=1)