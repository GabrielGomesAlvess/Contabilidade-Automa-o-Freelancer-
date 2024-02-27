'''
Site usado na automatização 
https://cadastro-produtos-devaprender.netlify.app/index.html
'''

import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook ['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(484,337, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(477,448, duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(553,556, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(529,642, duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(502,729, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(495,824, duration=1)
    pyautogui.hotkey('ctrl','v')

    #botão próximo
    pyautogui.click(342,877, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(433,364, duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(426,448, duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(428,528, duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(425,613, duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(386,703, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(356,733, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(352,757, duration=1)
    else:
        pyautogui.click(344,788, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(350,791, duration=1)
    pyautogui.hotkey('ctrl','v')

    #botão próximo
    pyautogui.click(344,853, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(395,375, duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(385,465, duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(385,572, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(376,685, duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(373,774, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão concluir
    pyautogui.click(355,829, duration=1)
    sleep(2)

    #Botão confirmar
    pyautogui.click(1135,188, duration=1)
    sleep(2)

    #Botão adicionar mais um
    pyautogui.click(947,609, duration=1)
    sleep(2)

