# Automatizar o envio de Mensagens
# "Debugar" - F9, F5, escolha Python File, Aba DEBUG CONSOLE e verificar a linha que está selecionada

# Importar Bibliotecas
import openpyxl                     # Abre o Arquivo Excel
from urllib.parse import quote      # Formata links especiais como no linkmensagem abaixo 
import webbrowser                   # Abre o seu navegador
from time import sleep              # Dar tempo para os comandos serem executados
import pyautogui                    # Automatizar o comando de envio de mensagem

webbrowser.open('https://web.whatsapp.com/')  # Acessar e dar tempo para logar a aplicação ao celular
sleep(10)

# Abrir o arquivo Excel com os dados
arquivo = openpyxl.load_workbook('TbTzMessages.xlsx')
pagina = arquivo['Sheet1']

# Criar um LOOP para puxar os dados de cada linha do arquivo
for linha in pagina.iter_rows(min_row=2):
    # Criar variáveis para puxar cada célula do arquivo
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = linha[2].value
    
    # Criar uma mensagem padrão
    mensagemwhatsapp = f'Olá, {nome}. {mensagem}'
    
    try:        # Criar um comando para salvar os contatos não enviados
        
        # Criar link personalizado com o Whatsapp Web
        # https://web.whatsapp.com/send?phone=xxxxx&text=xxxxx
        linkmensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagemwhatsapp)}'
    
        # Abrir aba do navegador no Whatsapp Web já logado
        webbrowser.open(linkmensagem)
        sleep(10)
    
        # Criar uma variável para encontrar a imagem para clicar e enviar a mensagem
        seta = pyautogui.locateCenterOnScreen('seta.png')  # Necessita pip install pillow
        sleep(5)

        # Comando para clicar no envio de mensagem onde o [0] e [1] refere-se ao local exato do botão no navegador
        pyautogui.click(seta[0], seta[1])
        sleep(5)

        # Comando para fechar o navegador sempre que enviar uma mensagem
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
        
    except:
        print(f'Não foi possível enviar para {nome}')
        
        # Comando para criar um arquivo com o não enviados onde: 'a' (append), newline='' (para incluir linhas), encoding='utf-8' (codificado para o Brasil)
        with open('LogErros.csv', 'a', newline='', encoding='utf-8') as arquivoerro:  
            arquivoerro.write(f'{nome},{telefone}')
    