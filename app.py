import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook(r'C:\Users\Daniel\Desktop\Projeto Freelancer\produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1518,305, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1472,413, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1467,526, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1481,614, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1463,691, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1465,783, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão para Mudar de Página
    pyautogui.click(1456,854, duration=1)
    sleep(4)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1539,325, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1515,411, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1504,501, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1489,574, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    tamanho = linha[10].value
    pyautogui.click(1530,670, duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(1520,705, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1509,730, duration=1)
    else:
        pyautogui.click(1526,743, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1496,754, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão para Mudar de Página
    pyautogui.click(1494,808, duration=1)
    sleep(4)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1465,347, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1473,426, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1475,523, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1471,655, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1472,733, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão para Concluir
    pyautogui.click(1485,814, duration=1)
    sleep(4)

    # Botão para Confirmar Inclusão
    pyautogui.click(1887,198, duration=1)
    sleep(4)

    # Botão para a Segunda Confirmação
    pyautogui.click(1886,191, duration=1)
    sleep(4)

    # Botão para a Iniciar Cadastro Novamente
    pyautogui.click(1695,580, duration=1)
    sleep(4)