# Rastreamento Automático de Pacotes dos Correios
Este é um script Python para automatizar o rastreamento de pacotes dos Correios utilizando a API pública do site linkcorreios.com.br.

# Pré-requisitos
Python 3.x instalado
Bibliotecas Python: pandas, requests, beautifulsoup4, openpyxl
Você pode instalar as bibliotecas necessárias executando o seguinte comando no terminal:

pip install pandas requests beautifulsoup4 openpyxl


# Como Usar
Clone este repositório ou baixe o arquivo rastreio.py.
Certifique-se de que sua planilha do Excel contém uma coluna com os códigos de rastreio dos pacotes. O nome da coluna deve ser especificado na variável coluna_rastreamento dentro do código.
Execute o script rastreio.py.
Aguarde até que o script termine de processar todas as linhas da planilha.
Após a execução, o script atualizará a planilha original com os dados de status, data de entrega e observações de status.
O tempo total de execução e o número total de pacotes atualizados serão exibidos como resultado no console.


# Personalização
Se desejar alterar o nome das colunas na planilha Excel, você pode modificar as variáveis coluna_status, coluna_data e coluna_obs_status no código.
O URL base do site de rastreamento dos Correios é definido como https://linkcorreios.com.br/. Se o site mudar de endereço, você pode atualizar a variável url_base no código.


# Autor
Este script foi desenvolvido por Gustavo Macena e está disponível sob a licença MIT.

