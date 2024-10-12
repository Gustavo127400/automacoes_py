import pyautogui
import time
import pandas as pd

link = 'https://dlp.hashtagtreinamentos.com/python/intensivao/tabela'

# Configuração padrão de tempo para pausas entre ações
pyautogui.PAUSE = 0.5

def entrar_site():
    # Abrir o navegador e acessar o link
    pyautogui.press('win')
    pyautogui.write('Edge')  # Corrigido o nome do navegador
    pyautogui.press('enter')
    time.sleep(1)  # Aguarda o navegador abrir
    pyautogui.click(x=1504, y=35)  # Clica na barra de endereços
    pyautogui.write(link)
    pyautogui.press('enter')
    time.sleep(3)  # Espera o site carregar completamente

    # Scroll para rolar a página para baixo
    pyautogui.scroll(-500)  # Rola para baixo (o valor negativo define a direção)

# Fazer Login
def entrada_email():
    # Clicar no campo de e-mail
    pyautogui.click(x=650, y=477) 
    pyautogui.write('gustavo@gmail.com')  # Entrada de dados - E-MAIL
    pyautogui.press('tab')
    
    # Entrada de dados - SENHA
    pyautogui.write('gustavo123')
    pyautogui.press('tab')
    
    # Confirmação
    pyautogui.press('enter')
    time.sleep(3)  # Aguarda o login ser processado

# Importar a base de dados
base_de_dados = pd.read_excel('df_produtos.xlsx')  # Certifique-se de que o arquivo está no caminho correto

# Cadastrar produtos
def cadastrar_produto():
    for linha in base_de_dados.index:
        pyautogui.click(x=801, y=308)  # Necessita das coordenadas
        time.sleep(0.5)  # Pausa para garantir que o sistema esteja pronto
        
        # Código do Produto
        codigo = base_de_dados.loc[linha, 'codigo']
        pyautogui.write(str(codigo))
        pyautogui.press('tab')

        # Marca do Produto
        marca = base_de_dados.loc[linha, 'marca']
        pyautogui.write(str(marca))
        pyautogui.press('tab')

        # Tipo do Produto
        tipo = base_de_dados.loc[linha, 'tipo']
        pyautogui.write(str(tipo))
        pyautogui.press('tab')

        # Categoria do Produto
        categoria = base_de_dados.loc[linha, 'categoria']
        pyautogui.write(str(categoria))
        pyautogui.press('tab')

        # Preço Unitário do Produto
        preco = base_de_dados.loc[linha, 'preco_unitario']
        pyautogui.write(str(preco))
        pyautogui.press('tab')

        # Custo do Produto
        custo = base_de_dados.loc[linha, 'custo']
        pyautogui.write(str(custo))
        pyautogui.press('tab')

        # Observação
        obs = base_de_dados.loc[linha, 'obs']
        pyautogui.write(str(obs))
        pyautogui.press('tab')

        # Confirmação dos dados
        pyautogui.press('enter')
        time.sleep(1)  # Pausa entre o cadastro de cada produto para processar

        # Rolar a tela se necessário para ver mais produtos
        pyautogui.scroll(1000)  # Rola a tela para baixo

# Execução das funções
entrar_site()  # Entrar no site
cadastrar_produto()  # Cadastrar os produtos
