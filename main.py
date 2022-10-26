# Importa o tkinter com o apelido tk
import tkinter as tk

# Importa o módulo ttk do tkinter
from tkinter import StringVar, ttk

# Importa a classe DateEntry
from tkcalendar import DateEntry

# Importa a biblioteca requests
import requests

# Importa a função para abrir arquivos
from tkinter.filedialog import askopenfilename

# Importa o pandas com o apelido pd
import pandas as pd

# Importa o numpy com o apelido np'
import numpy as np


def obter_moedas_disponiveis():
    # Link da requisição
    link = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/Moedas?$top=100&$format=json"

    # Resposta da requisição para obter todas as moedas
    resposta = requests.get(link)

    # Dicionário com os dados da requisição
    dados_requisicao = resposta.json()

    # Lista de dicionários com os dados da requisição
    moedas = dados_requisicao["value"]

    # Lista de moedas disponíveis na API
    moedas_disponiveis = [moeda["simbolo"] for moeda in moedas]

    # Retorna uma lista com as moedas disponíveis
    return moedas_disponiveis


def converter_formato_data(data):
    # Extrai o dia da data
    dia = data[0:2]

    # Extai o mês da data
    mes = data[3:5]

    # Extrai o ano da data
    ano = data[6:]

    return f"{mes}-{dia}-{ano}"


def extrair_data(carimbo_data_hora):
    # Extrai o dia do carimbo de data/hora
    dia = carimbo_data_hora[8:10]

    # Extrai o mês do carimbo de data/hora
    mes = carimbo_data_hora[5:7]

    # Extrai o ano do carimbo de data/hora
    ano = carimbo_data_hora[:4]

    return f"{dia}/{mes}/{ano}"


def obter_cotacao():
    # Moeda selecionada
    moeda_selecionada = menu_moedas.get()

    # Data selecionada
    data_selecionada = campo_data.get()

    # Data formatada
    data_selecionada = converter_formato_data(data_selecionada)

    # Link para fazer a requisição
    link = f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/" \
           f"CotacaoMoedaDia(moeda=@moeda,dataCotacao=@dataCotacao)" \
           f"?@moeda='{moeda_selecionada}'&@dataCotacao='{data_selecionada}'" \
           f"&$top=1&$format=json&$select=cotacaoCompra"

    # Resposta da requisição
    resposta = requests.get(link)

    # Dados da requisição
    dados_requisicao = resposta.json()

    # Cotação da moeda selecionada
    cotacao = dados_requisicao["value"][0]["cotacaoCompra"]

    # Atualiza a mensagem de cotação
    rotulo_cotacao["text"] = f"Cotação do {moeda_selecionada}: R$ {cotacao}"


def selecionar_planilha():
    # Armazena o caminho da planilha selecionada
    caminho = askopenfilename(title="Selecione uma planilha com as moedas")

    # Redefine o valor da variável que armazena o caminho do arquivo selecionado
    caminho_arquivo_selecionado.set(caminho)

    if caminho:
        # Atualiza a mensagem de arquivo selecionado
        mensagem_caminho_arquivo["text"] = f"Arquivo selecionado: {caminho}"


def obter_cotacoes():
    # Importa a planilha selecionado como um dataframe do pandas
    df_moedas = pd.read_excel(caminho_arquivo_selecionado.get())

    # Obtém a lista de moedas da coluna A
    moedas = df_moedas.iloc[:, 0]

    # Data inicial
    data_inicial = converter_formato_data(campo_data_inicial.get())

    # Data final
    data_final = converter_formato_data(campo_data_final.get())

    # Percorre a lista de moedas
    for moeda in moedas:
        # Link da requisição
        link = f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/" \
               f"CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial," \
               f"dataFinalCotacao=@dataFinalCotacao)?@moeda='{moeda}'&@dataInicial='{data_inicial}'" \
               f"&@dataFinalCotacao='{data_final}'&$top=100&$format=json" \
               f"&$select=cotacaoCompra,dataHoraCotacao"

        # Resposta da requisição
        resposta = requests.get(link)

        # Dados da requisição (dicionário)
        dados_requisicao = resposta.json()

        # Lista de cotações
        cotacoes = dados_requisicao["value"]

        # Pecorre a lista de cotações
        for cotacao in cotacoes:
            # Valor da cotação
            valor = cotacao["cotacaoCompra"]

            # Carimbo de data/hora da cotação
            carimbo_data_hora = cotacao["dataHoraCotacao"]

            # Converte o carimbo de data/hora numa data do tipo dd/mm/aaaa
            data = extrair_data(carimbo_data_hora)

            # Verifica se a coluna de data está presente no dataframe
            if data not in df_moedas:
                # Cria a coluna com a data no cabeçalho e preenche a célula com NaN
                df_moedas[data] = np.nan

            # Preenche a célula com o valor da cotação
            df_moedas.loc[df_moedas.iloc[:, 0] == moeda, data] = valor

    # Exporta o dataframe como uma planilha do Excel
    df_moedas.to_excel("cotacao-moesas.xlsx")

    # Atualiza a mensagem de cotações
    rotulo_mensagem_atualizacao["text"] = "Cotações obtidas com sucesso"


# Lista de moedas disponíveis
moedas_disponiveis = obter_moedas_disponiveis()

# Cria uma janela
janela = tk.Tk()

# Faz as colunas se ajustarem a janela
janela.columnconfigure([0, 1, 2], weight=1)

# Altera o ícone da janela
janela.iconbitmap("./assets/favicon.ico")

# Modifica o título da janela
janela.title("Ferramenta de Cotação de Moedas")

# Obtenção das dimensões da tela
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()

# Definição das dimensões da janela
largura_janela = 600
altura_janela = 480

# Coordenadas do centro da tela
centro_x = int(largura_tela / 2 - largura_janela)
centro_y = int(altura_tela / 2 - altura_janela / 2)

# Centraliza a janela na tela
janela.geometry(f"{largura_janela}x{altura_janela}+{centro_x}+{centro_y}")

# Rótulo do primeiro cabeçalho da janela
rotulo_cotacao_moeda = ttk.Label(
    text="Cotação de Moeda Específica",
    borderwidth=2,
    relief="solid",
    anchor="center"
)

# Posiciona o rótulo do cabeçalho
rotulo_cotacao_moeda.grid(
    row=0,
    column=0,
    columnspan=3,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo para caixa de seleção de moeda
rotulo_menu_moedas = ttk.Label(
    text="Selecione uma moeda",
    anchor="w"
)

# Posiciona a caixa de seleção de moeda
rotulo_menu_moedas.grid(
    row=1,
    column=0,
    columnspan=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Caixa de seleção de moedas
menu_moedas = ttk.Combobox(values=moedas_disponiveis)

# Posiciona a caixa de seleção de moedas
menu_moedas.grid(
    row=1,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo para caixa de seleção de data
rotulo_data_cotacao = ttk.Label(
    text="Selecione a data da cotação",
    anchor="w"
)

# Posiciona o rótulo de data da cotação da moeda
rotulo_data_cotacao.grid(
    row=2,
    column=0,
    columnspan=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Menu de seleção de data
campo_data = DateEntry(year=2022, locale="pt_br")

# Posiciona o menu de seleção de data
campo_data.grid(
    row=2,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo com a mensagem retornando a cotação
rotulo_cotacao = ttk.Label(
    text="",
    anchor="w"
)

# Posiciona a mensagem de cotação
rotulo_cotacao.grid(
    row=3,
    column=0,
    columnspan=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Botão para obter a cotação
botao_obter_cotacao = ttk.Button(
    text="Obter Cotação",
    command=obter_cotacao
)

# Posiciona o botão para obter a cotação
botao_obter_cotacao.grid(
    row=3,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo do segundo cabeçalho da janela
rotulo_cotacao_moedas = ttk.Label(
    text="Cotação de Moedas Específicas",
    borderwidth=2,
    relief="solid",
    anchor="center"
)

# Posiciona o rótulo na janela
rotulo_cotacao_moedas.grid(
    row=4,
    column=0,
    columnspan=3,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo para campo de seleção de arquivo
rotulo_selecao_arquivo = ttk.Label(
    text="Selecione uma planilha com os símbolos das moedas na coluna A",
    anchor="w"
)

# Posiciona o rótulo para seleção de arquivo
rotulo_selecao_arquivo.grid(
    row=5,
    column=0,
    columnspan=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Cria uma variável do tkinter
caminho_arquivo_selecionado = StringVar()

# Campo de seleção de arquivo
botao_selecao_arquivo = ttk.Button(
    text="Selecionar Planilha",
    command=selecionar_planilha
)

# Posiciona o campo de seleção de arquivo
botao_selecao_arquivo.grid(
    row=5,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Mensagem com caminho do arquivo selecionado
mensagem_caminho_arquivo = ttk.Label(
    text="Nenhum arquivo selecionado: ",
    anchor='e'
)

# Posiciona a mensagem do camino do arquivo selecionado
mensagem_caminho_arquivo.grid(
    row=6,
    column=0,
    columnspan=3,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo de data inicial
rotulo_data_inicial = ttk.Label(
    text="Data Inicial",
    anchor="w"
)

# Posiciona o rótulo de data inicial
rotulo_data_inicial.grid(
    row=7,
    column=0,
    sticky="NESW",
    padx=10,
    pady=10
)

# Campo de data inicial
campo_data_inicial = DateEntry(year=2022, locale="pt_br")

# Posiciona o campo de data inicial
campo_data_inicial.grid(
    row=7,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo de data final
rotulo_data_final = ttk.Label(
    text="Data Final",
    anchor="w"
)

# Posiciona o rótula de data final
rotulo_data_final.grid(
    row=8,
    column=0,
    sticky="NESW",
    padx=10,
    pady=10
)

# Campo de data final
campo_data_final = DateEntry(year=2022, locale="pt_br")

# Posiciona o campo de data final
campo_data_final.grid(
    row=8,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Botão de atualizar cotações
botao_atualizar_cotacoes = ttk.Button(
    text="Atualizar Cotações",
    command=obter_cotacoes
)

# Posiciona o botão de atualizar cotações
botao_atualizar_cotacoes.grid(
    row=9,
    column=0,
    sticky="NESW",
    padx=10,
    pady=10
)

# Rótulo da mensagem de atualização da planilha
rotulo_mensagem_atualizacao = ttk.Label(text="")

# Posiciona o rótulo da mensagem de atualização da planilha
rotulo_mensagem_atualizacao.grid(
    row=9,
    column=1,
    columnspan=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Botão fechar
botao_fechar = ttk.Button(text="Fechar", command=janela.quit)

# Posiciona o botão fechar
botao_fechar.grid(
    row=10,
    column=2,
    sticky="NESW",
    padx=10,
    pady=10
)

# Coloca a janela em loop
janela.mainloop()
