import openpyxl  # Biblioteca para manipular arquivos Excel (.xlsx)
import pyperclip  # Biblioteca para copiar textos para a área de transferência
import pyautogui  # Biblioteca para automação de ações no computador (cliques, teclas, etc.)
from time import sleep  # Biblioteca para adicionar pausas no código

# Carrega a planilha Excel com os dados dos produtos
workbook = openpyxl.load_workbook(r'C:\Users\Daniel\Desktop\Projeto Freelancer\produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']  # Seleciona a aba "Produtos"

# Percorre cada linha da aba, a partir da segunda (para ignorar o cabeçalho)
for linha in sheet_produtos.iter_rows(min_row=2):
    # Obtém o nome do produto e insere no campo correspondente
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)  # Copia o nome do produto para a área de transferência
    pyautogui.click(1518, 305, duration=1)  # Clica no campo de "Nome do Produto"
    pyautogui.hotkey('ctrl', 'v')  # Cola o conteúdo no campo

    # Repete o processo para outros campos da planilha
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1472, 413, duration=1)  # Campo de "Descrição"
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1467, 526, duration=1)  # Campo de "Categoria"
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1481, 614, duration=1)  # Campo de "Código do Produto"
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1463, 691, duration=1)  # Campo de "Peso"
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1465, 783, duration=1)  # Campo de "Dimensões"
    pyautogui.hotkey('ctrl', 'v')

    # Avança para a próxima página
    pyautogui.click(1456, 854, duration=1)
    sleep(4)

    # Continua preenchendo outros campos do formulário
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1539, 325, duration=1)  # Campo de "Preço"
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1515, 411, duration=1)  # Campo de "Quantidade em Estoque"
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1504, 501, duration=1)  # Campo de "Data de Validade"
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1489, 574, duration=1)  # Campo de "Cor"
    pyautogui.hotkey('ctrl', 'v')

    # Preenchimento do campo "Tamanho"
    tamanho = linha[10].value
    pyautogui.click(1530, 670, duration=1)  # Abre opções de tamanho
    if tamanho == 'Pequeno':
        pyautogui.click(1520, 705, duration=1)  # Seleciona "Pequeno"
    elif tamanho == 'Médio':
        pyautogui.click(1509, 730, duration=1)  # Seleciona "Médio"
    else:
        pyautogui.click(1526, 743, duration=1)  # Seleciona "Grande"

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1496, 754, duration=1)  # Campo de "Material"
    pyautogui.hotkey('ctrl', 'v')

    # Avança para a próxima página
    pyautogui.click(1494, 808, duration=1)
    sleep(4)

    # Preenchimento de campos adicionais
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1465, 347, duration=1)  # Campo de "Fabricante"
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1473, 426, duration=1)  # Campo de "País de Origem"
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1475, 523, duration=1)  # Campo de "Observações"
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1471, 655, duration=1)  # Campo de "Código de Barras"
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1472, 733, duration=1)  # Campo de "Localização no Armazém"
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