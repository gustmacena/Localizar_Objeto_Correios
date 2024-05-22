# Importar as bibliotecas necessárias
import pandas as pd             # Para trabalhar com dataframes e arquivos do Excel
import requests                 # Para fazer solicitações HTTP
from bs4 import BeautifulSoup  # Para fazer scraping de dados HTML
from openpyxl import load_workbook  # Para carregar e salvar arquivos Excel
import time                     # Para medir o tempo de execução
import threading                # Para criar threads
import itertools                # Para criar animação de carregamento

# Função para exibir "Loading ..." piscando no terminal
def loading_animation(stop_event):
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if stop_event.is_set():
            break
        print(f'\rLoading {c}', end='', flush=True)
        time.sleep(0.1)

# Iniciar a contagem do tempo para medir a duração da execução
inicio_tempo = time.time()

# Carregar a planilha do Excel que contém os códigos de rastreio
excel_file = 'Ecommerce_Api_Correios.xlsx'  # Nome do arquivo Excel
dados = pd.read_excel(excel_file)           # Carrega o arquivo Excel em um DataFrame do pandas

# Definir os nomes das colunas que contêm os dados de rastreio e as informações de retorno
coluna_rastreamento = 'RASTREIO'         # Nome da coluna que contém os códigos de rastreio
coluna_status = 'STATUS'                 # Nome da coluna para armazenar o status do rastreio
coluna_data = 'DATA'                     # Nome da coluna para armazenar a data de entrega
coluna_obs_status = 'OBS STATUS'         # Nome da coluna para observações adicionais de status

# Garantir que as colunas onde vamos atualizar os dados estejam no tipo de dado correto
dados[coluna_status] = dados[coluna_status].astype(str)
dados[coluna_data] = dados[coluna_data].astype(str)
dados[coluna_obs_status] = dados[coluna_obs_status].astype(str)

# URL base do site de rastreamento dos Correios
url_base = 'https://linkcorreios.com.br/'

# Função para buscar o status, data e local do rastreio com base no código de rastreio
def buscar_rastreamento(codigo_rastreamento):
    # Montar a URL completa para o código de rastreio fornecido
    url = f"{url_base}{codigo_rastreamento}"
    
    # Enviar uma solicitação HTTP GET para a URL e obter a resposta
    response = requests.get(url)

    # Verificar se a solicitação foi bem-sucedida (código de status 200)
    if response.status_code == 200:
        # Analisar o conteúdo HTML da página usando BeautifulSoup
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Tentar encontrar os elementos HTML que contêm o status, data/hora e local do rastreio
        try:
            status_element = soup.select_one('ul.linha_status li:-soup-contains("Status:")')
            data_hora_element = soup.select_one('ul.linha_status li:-soup-contains("Data  :")')
            local_element = soup.select_one('ul.linha_status li:-soup-contains("Local:")')
            
            # Se todos os elementos forem encontrados, extrair as informações
            if status_element and data_hora_element and local_element:
                status = status_element.text.split(':', 1)[1].strip()  # Extrair o status
                data_hora = data_hora_element.text.split(':', 1)[1].strip()  # Extrair data/hora
                data = f"{data_hora.split('|')[0].strip()} {data_hora.split('|')[1].replace('Hora:', '').strip()}"
                # Extrair a data e a hora, formatadas como "dd/mm/aaaa hh:mm"
                local = local_element.text.split(':', 1)[1].strip()  # Extrair o local
                obs_status = local  # Armazenar o local como observação de status
            else:
                # Se algum elemento estiver faltando, definir status e data como "Não encontrado"
                status = "Não encontrado"
                data = ""
                obs_status = ""
        except AttributeError:
            # Se ocorrer um erro ao procurar elementos, definir status e data como "Não encontrado"
            status = "Não encontrado"
            data = ""
            obs_status = ""
    else:
        # Se a solicitação falhar, definir status e data como "Erro na conexão"
        status = "Erro na conexão"
        data = ""
        obs_status = ""

    return status, data, obs_status  # Retornar o status, data e observação de status

# Criar e iniciar a thread de animação de carregamento
stop_event = threading.Event()
loading_thread = threading.Thread(target=loading_animation, args=(stop_event,))
loading_thread.start()

# Percorrer cada linha da planilha para buscar e atualizar as informações de rastreio
for index, linha in dados.iterrows():
    codigo_rastreamento = linha[coluna_rastreamento]  # Obter o código de rastreio da linha atual
    
    # Chamar a função buscar_rastreamento para obter o status, data e observação de status
    status, data, obs_status = buscar_rastreamento(codigo_rastreamento)
    
    # Atualizar as colunas de status, data e observação de status na linha atual
    dados.loc[index, coluna_status] = status
    dados.loc[index, coluna_data] = data
    dados.loc[index, coluna_obs_status] = obs_status

# Carregar a planilha original com openpyxl para preservar a formatação
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    sheet_name = writer.book.sheetnames[0]  # Obter o nome da primeira planilha no arquivo Excel
    dados.to_excel(writer, index=False, sheet_name=sheet_name)  # Salvar os dados atualizados na planilha

# Parar a thread de animação de carregamento
stop_event.set()
loading_thread.join()

# Calcular o tempo total de execução em minutos e segundos
fim_tempo = time.time()
tempo_total = fim_tempo - inicio_tempo
minutos = int(tempo_total // 60)
segundos = int(tempo_total % 60)

# Limpar a linha de "Loading ..."
print('\r' + ' ' * 20 + '\r', end='')

# Imprimir um resumo da busca, incluindo o número total de pacotes atualizados e o tempo total de execução em minutos e segundos
print(f"Busca de rastreamento concluída! {len(dados)} pacotes atualizados em {minutos} minutos e {segundos} segundos.")
