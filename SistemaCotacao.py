import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import requests
import json

# --- Requisição da Api ---

try:
    pares_comuns = "USD-BRL,EUR-BRL,BTC-BRL,GBP-BRL,JPY-BRL"
    # Pegar todas os pares de cotações
    requisicao = requests.get(f'https://economia.awesomeapi.com.br/json/last/{pares_comuns}')
    requisicao.raise_for_status() # Encontrar um HTTPError na resposta
    # Tranformar a requisição Json em dicionário Python
    dicionario_moedas = requisicao.json()

    # Criar uma lista com as chaves do dicionário
    lista_moedas = sorted(list(set([chave[:3] for chave in dicionario_moedas.keys()])))

except requests.exceptions.RequestException as e:
    print(f'Erro nos dados da API: {e}')
    lista_moedas = []


# --- Funções do Aplicativo ---

# Modificar a função para aceitar o label como um argumento
def pegar_cotacao(label_feedback):
    '''
    Obter a cotação para uma moeda e data específica.
    '''
    # Setar um texto inicial no label que foi passado
    label_feedback['text'] = "Buscando cotação..."

    selecionar_moeda = combobox_moeda.get()
    selecionar_data_cotacao = calendario_moeda.get()

    if not selecionar_moeda:
        label_feedback['text'] = 'Por favor, selecione uma moeda.'
        return

    if not selecionar_data_cotacao:
        label_feedback['text'] = 'Por favor, selecione uma data.'
        return

    try:        
        objeto_data_cotacao = datetime.strptime(selecionar_data_cotacao, '%d/%m/%Y') # Formato de data obtido
        formato_api_data = datetime.strftime(objeto_data_cotacao, '%Y%m%d') # Formato de data para a API

        link = f'https://economia.awesomeapi.com.br/json/daily/{selecionar_moeda}-BRL/?start_date={formato_api_data}&end_date={formato_api_data}'

        requisicao_moeda = requests.get(link)
        requisicao_moeda.raise_for_status()
        data_cotacao = requisicao_moeda.json()

        if data_cotacao and isinstance(data_cotacao, list) and len(data_cotacao) > 0:
            valor_bid = float(data_cotacao[0]['bid'])
            label_feedback['text'] = f'A cotação da moeda {selecionar_moeda} na data {selecionar_data_cotacao} foi de R$ {valor_bid:.4f}.'
        else:
            label_feedback['text'] = f'Nenhuma cotação encontrada para {selecionar_moeda} na {selecionar_data_cotacao}.'

    except requests.exceptions.RequestException as e:
        label_feedback['text'] = f'Erro de conexão ao buscar cotação: {e}'
    except ValueError:
        label_feedback['text'] = 'Formato de data inválido. Por favor, selecione novamente.'
    except IndexError:
        label_feedback['text'] = 'A cotação não foi encontrada para a seleção da moeda/data.'
    except Exception as e:
        label_feedback['text'] = f'Um erro inesperado ocorreu: {e}'


def selecionar_arquivo():
    '''
    Abrir uma pasta para selecionar um arquivo Excel.
    '''
    caminho_arquivo = askopenfilename(title='Selecionar a moeda no arquivo Excel', filetypes=[("Excel files", "*.xlsx *.xls")])
    var_caminho_arquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivo_selecionado['text'] = f"Arquivo selecionado: {caminho_arquivo.split('/')[-1]}"
    else:
        label_arquivo_selecionado['text'] = "Nenhum arquivo selecionado."


def atualizar_cotacoes():
    '''
    Ler um arquivo Excel com moedas códigos e updates com histórico de cotações.
    '''
    caminho_arquivo = var_caminho_arquivo.get()
    if not caminho_arquivo:
        label_cotacoes_atualizadas['text'] = 'Por favor, selecione primeiro um arquivo Excel.'
        return

    try:
        df = pd.read_excel(caminho_arquivo)
        if df.empty or df.iloc[:, 0].empty:
            label_cotacoes_atualizadas['text'] = 'A seleção do arquivo Excel está vazia ou não há moedas na primeira coluna.'
            return
        
        moedas_excel = df.iloc[:, 0].dropna().unique().tolist() # Selecionar um único tipo de moeda

        data_inicial_str = calendario_data_inicial.get()
        data_final_str = calendario_data_final.get()

        data_inicial = datetime.strptime(data_inicial_str, '%d/%m/%Y') # Converter para o formato brasileiro
        data_final = datetime.strptime(data_final_str, '%d/%m/%Y') # Converter para o formato brasileiro

        if data_inicial > data_final:
            label_cotacoes_atualizadas['text'] = 'A data inicial não pode ser depois da data final.'
            return

        delta = data_final - data_inicial
        update_df = df.copy()

        for moeda in moedas_excel:
            api_data_inicial = data_inicial.strftime('%Y%m%d')
            api_data_final = data_final.strftime('%Y%m%d')

            link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={api_data_inicial}&end_date={api_data_final}'

            try:
                requisicao_moeda = requests.get(link)
                requisicao_moeda.raise_for_status()
                cotacoes = requisicao_moeda.json()

                if cotacoes and isinstance(cotacoes, list):
                    for cotacao in cotacoes:
                        if 'timestamp' in cotacao and 'bid' in cotacao:
                            timestamp_dt = datetime.fromtimestamp(int(cotacao['timestamp']))
                            data_formatada = timestamp_dt.strftime('%d/%m/%Y')
                            bid = float(cotacao['bid'])

                            if data_formatada not in update_df.columns:
                                update_df[data_formatada] = np.nan

                            update_df.loc[update_df.iloc[:, 0] == moeda, data_formatada] = bid
                        else:
                            print(f"Dados incompletos para {moeda} no período. Pulando esta cotação.")
                else:
                    print(f'Sem cotações na data retornada para {moeda} entre {data_inicial_str} e {data_final_str}.')
            except requests.exceptions.RequestException as e:
                print(f"Erro ao carregar dados para {moeda} (API/Rede): {e}")
            except (KeyError, ValueError) as e:
                print(f"Erro ao processar dados da API para {moeda} (Chave/Valor inválido): {e}")
            except Exception as e:
                print(f"Um erro inesperado ocorreu no processo para {moeda}: {e}")

        saida_caminho_arquivo = caminho_arquivo.replace('.xlsx', '_updated.xlsx').replace('.xls', '_updated.xls')
        update_df.to_excel(saida_caminho_arquivo, index=False)
        label_cotacoes_atualizadas['text'] = f'Cotações atualizadas e salvas em: {saida_caminho_arquivo}'

    except FileNotFoundError:
        label_cotacoes_atualizadas['text'] = 'Erro: Arquivo não encontrado. Por favor, selecione um arquivo Excel válido.'
    except pd.errors.EmptyDataError:
        label_cotacoes_atualizadas['text'] = 'Erro: O arquivo Excel selecionado está vazio.'
    except Exception as e:
        label_cotacoes_atualizadas['text'] = f'Um erro ocorreu durante o update: {e}. Verifique o formato do arquivo Excel.'


# --- Criação da interface gráfica ---

janela = tk.Tk()
janela.title('Ferramenta de Cotação de Moedas')

# Criar uma janela responsiva
for i in range(3):
    janela.columnconfigure(i, weight=1)
for i in range(11):
    janela.rowconfigure(i, weight=1)

# --- Cotação de 1 moeda específica ---

# Criar labels
label_cotacao_moeda = tk.Label(janela, text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
label_cotacao_moeda.grid(row=0, column=0, padx=10, pady=10, sticky='NSWE', columnspan=3)

label_selecionar_moeda = tk.Label(janela, text='Selecione a moeda que desejar consultar:', anchor='e')
label_selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, sticky='NSWE',columnspan=2)

# Criar lista suspensa com as moedas
combobox_moeda = ttk.Combobox(janela, values=lista_moedas, state='readonly')
combobox_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='NSWE')
if lista_moedas:
    combobox_moeda.set(lista_moedas[0])

# Criar label para selecionar a data
label_selecionar_dia = tk.Label(janela, text='Selecione o dia que deseja pegar a cotação:', anchor='e')
label_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, sticky='NSEW', columnspan=2)

# Setar uma data inicial
calendario_moeda = DateEntry(janela, year=datetime.now().year, month=datetime.now().month, day=datetime.now().day, locale='pt_br', date_pattern='dd/mm/yyyy', mindate=None, maxdate=datetime.now())
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='NSWE')

label_texto_cotacao = tk.Label(janela, text='')
label_texto_cotacao.grid(row=3, column=0, padx=10, pady=10, sticky='NSWE', columnspan=2)

# Criar botão para pegar a cotação
botao_cotacao = tk.Button(janela, text='Pegar Cotação', command=lambda: pegar_cotacao(label_texto_cotacao))
botao_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='NSWE')

# --- Cotação de múltiplas moedas ---

label_cotacao_multiplas_moedas = tk.Label(janela, text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
label_cotacao_multiplas_moedas.grid(row=4, column=0, padx=10, pady=10, columnspan=3, sticky='NSEW')

label_selecionar_arquivo = tk.Label(janela, text='Selecione um arquivo em Excel com as Moedas na Coluna A:')
label_selecionar_arquivo.grid(row=5, column=0, padx=10, pady=10, sticky='NSWE', columnspan=2)

var_caminho_arquivo = tk.StringVar()

botao_selecionar_arquivo = tk.Button(janela, text='Clique para selecionar', command=selecionar_arquivo)
botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='NSWE')

label_arquivo_selecionado = tk.Label(janela, text='Nenhum arquivo selecionado', anchor='e')
label_arquivo_selecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='NSWE')

label_data_inicial = tk.Label(janela, text='Data Inicial:', anchor='e')
label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky='NSWE')

# Set initial date to today, but also set maxdate to today to avoid future dates by default,
# and mindate to an arbitrary old date to facilitate selection of past dates.
calendario_data_inicial = DateEntry(janela, year=datetime.now().year, month=datetime.now().month, day=datetime.now().day, locale='pt_br', date_pattern='dd/mm/yyyy', mindate=datetime(2000, 1, 1), maxdate=datetime.now())
calendario_data_inicial.grid(row=7, column=1, padx=10, pady=10, sticky='NSWE')

label_data_final = tk.Label(janela, text='Data Final:', anchor='e')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='NSWE')

# Similar setup for the end date calendar
calendario_data_final = DateEntry(janela, year=datetime.now().year, month=datetime.now().month, day=datetime.now().day, locale='pt_br', date_pattern='dd/mm/yyyy', mindate=datetime(2000, 1, 1), maxdate=datetime.now())
calendario_data_final.grid(row=8, column=1, padx=10, pady=10, sticky='NSWE')

botao_atualizar_cotacoes = tk.Button(janela, text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='NSWE')

label_cotacoes_atualizadas = tk.Label(janela, text='', anchor='w')
label_cotacoes_atualizadas.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='NSWE')

botao_fechar = tk.Button(janela, text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='NSWE')

janela.mainloop()