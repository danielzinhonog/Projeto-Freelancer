# Projeto-Freelancer

Bot Assistente para Lançamentos Contábeis
  
Este projeto foi desenvolvido para automatizar o processo de lançamento contábil em um sistema corporativo que, originalmente, era repetitivo e demandava muito tempo. A solução criada permite que o usuário organize as informações em uma planilha Excel e, com base nisso, o bot realiza os lançamentos de forma autônoma, enquanto o usuário pode se concentrar em outras atividades.


Principais Funcionalidades

Automação de Lançamentos: O bot lê os dados da planilha Excel utilizando a biblioteca openpyxl e preenche automaticamente os campos do sistema contábil com a ajuda de pyautogui.

Copiar e Colar Preciso: A biblioteca pyperclip é utilizada para garantir a transferência fiel dos dados da planilha para o sistema.

Interação Automatizada: Simula cliques e interações no sistema com precisão, reduzindo erros humanos.

Flexível e Adaptável: O script pode ser ajustado para atender diferentes formatos de planilha e fluxos de trabalho contábeis.


Tecnologias Utilizadas

Python
openpyxl para manipulação da planilha Excel.
pyperclip para copiar e colar informações.
pyautogui para automação das interações no sistema.

Planilha Excel: Utilizada como base para o input dos dados.


Como Funciona

O usuário insere os dados necessários em uma planilha Excel com colunas específicas.

O bot lê os dados da planilha linha por linha e preenche os campos correspondentes no sistema.

Após o preenchimento, o bot realiza confirmações e reinicia o processo até completar todos os lançamentos.


Uso

Este projeto é ideal para empresas que desejam automatizar tarefas contábeis repetitivas e aumentar a produtividade no ambiente de trabalho.
