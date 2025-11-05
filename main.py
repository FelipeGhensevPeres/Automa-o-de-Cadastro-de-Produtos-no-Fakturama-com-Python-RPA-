import pandas as pd
import pyautogui
from time import sleep

df_produtos = pd.read_excel(r"Produtos.xlsx")
df_produtos.loc[0,'Nome'] = 'Camera'
df_produtos.loc[0,'Descrição'] = 'Camera 1080p Logitech USB'

#ABRIR O FAKTURAMA
pyautogui.press('win')
sleep(0.5)
pyautogui.write('Fakturama.exe')
sleep(0.5)
pyautogui.press('enter')


sleep(15)

for i, linha in df_produtos.iterrows():
    #CLICAR EM NEW PRODUCT
    new_product = pyautogui.locateOnScreen(r'imagens esperar/novo_produto.png',grayscale=True,confidence=0.8)
    if new_product:
        pyautogui.click(new_product)
    else:
        sleep(0.5)
    sleep(2)


        #REGISTRAR O ID DO PRODUTO
    id_produto = pyautogui.locateOnScreen(r'imagens esperar/id_produto.png',grayscale=True,confidence=0.8)
    if id_produto:
        x = id_produto.left + 95
        y = id_produto.top + 10

        pyautogui.click(x,y)
        sleep(1)
        pyautogui.write(str(linha['ID']))
        sleep(0.5)


        #NOME DO PRODUTO
    nome_produto = pyautogui.locateOnScreen(r'imagens esperar/nome_produto.png',grayscale=True,confidence=0.8)
    if nome_produto:
        x = nome_produto.left + 95
        y = nome_produto.top + 10

        pyautogui.click(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Nome']))
        sleep(0.5)

            
    #CATEGORIA DO PRODUTO
    categoria_produto = pyautogui.locateOnScreen(r'imagens esperar/categoria.png',grayscale=True,confidence=0.8)
    if categoria_produto:
        x = categoria_produto.left + 95
        y = categoria_produto.top + 10

        pyautogui.click(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Categoria']))
        sleep(0.5)
            

    #GTIN DO PRODUTO
    gtin_produto = pyautogui.locateOnScreen(r'imagens esperar/GTIN.png',grayscale=True,confidence=0.8)
    if gtin_produto:
        x = gtin_produto.left + 95
        y = gtin_produto.top + 10

        pyautogui.click(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['GTIN']))
        sleep(0.5)

    #SUPLIER DO PRODUTO
    suplier_produto = pyautogui.locateOnScreen(r'imagens esperar/suplier.png',grayscale=True,confidence=0.8)
    if suplier_produto:
        x = suplier_produto.left + 95
        y = suplier_produto.top + 10

        pyautogui.click(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Supplier']))
        sleep(0.5)

    #DESCRICAO DO PRODUTO
    descricao_produto = pyautogui.locateOnScreen(r'imagens esperar/descricao.png',grayscale=True,confidence=0.8)
    if descricao_produto:
        x = descricao_produto.left + 95
        y = descricao_produto.top + 10

        pyautogui.click(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Descrição']))
        sleep(0.5)

            
    #PRECO DO PRODUTO
    preco_produto = pyautogui.locateOnScreen(r'imagens esperar/preco.png',grayscale=True,confidence=0.8)
    if preco_produto:
        x = preco_produto.left + 110
        y = preco_produto.top + 10
        pyautogui.doubleClick(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Preço']))
        sleep(0.5)
            
    #CUSTO DO PRODUTO
    custo_produto = pyautogui.locateOnScreen(r'imagens esperar/custo.png',grayscale=True,confidence=0.8)
    if custo_produto:
        x = custo_produto.left + 100
        y = custo_produto.top + 10
        pyautogui.doubleClick(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Custo']))
        sleep(0.5)
            
    #ESTOQUE DO PRODUTO
    estoque_produto = pyautogui.locateOnScreen(r'imagens esperar/estoque.png',grayscale=True,confidence=0.8)
    if estoque_produto:
        x = estoque_produto.left + 100
        y = estoque_produto.top + 10
        pyautogui.doubleClick(x,y)
        sleep(0.5)
        pyautogui.write(str(linha['Estoque']))
        sleep(0.5)
    
    #COLOCAR A IMAGEM DO PRODUTO
    selecionar_imagem = pyautogui.locateOnScreen(r'imagens esperar/imagem.png',grayscale=True,confidence=0.8)
    if selecionar_imagem:
        pyautogui.click(selecionar_imagem)
        sleep(0.5)
        pyautogui.write(rf"C:\Users\Master\Videos\ESTUDOS HASHTAG\PYTHON\31 - RPA com Python Automacao de ERPs\Imagens Produtos\{linha['Imagem']}")
        pyautogui.press('enter')
        sleep(2)

    #ADICIONAR O PRODUTO COM O SAVE

    salvar_produto = pyautogui.locateOnScreen(r'imagens esperar/salvar.png',grayscale=True,confidence=0.8)
    if salvar_produto:
        pyautogui.click(salvar_produto)
        sleep(0.5)