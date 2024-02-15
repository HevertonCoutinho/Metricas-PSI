import requests
import csv
import datetime
from openpyxl import Workbook
from tqdm import tqdm


def get_pagespeed_metrics(urls):
    # Solicitar chave da API
    api_key = input('Digite sua chave da API: ')
        
    metrics_list = []

    # Parâmetros da API
    params = {
        "strategy": "mobile",
        "category": "performance",
        "locale": "pt_BR",
        "key": api_key
    }

    for url in tqdm(urls, desc='Processando', unit='URL'):
        params["url"] = url

        try:
            response = requests.get("https://www.googleapis.com/pagespeedonline/v5/runPagespeed", params=params)
            data = response.json()

            if 'lighthouseResult' in data:
                metrics = {}
                metrics['url'] = url
                metrics['score'] = data['lighthouseResult']['categories']['performance']['score']

                audits = data['lighthouseResult']['audits']

                metrics_to_check = [
                    ('fcp', 'first-contentful-paint'),
                    ('lcp', 'largest-contentful-paint'),
                    ('tti', 'interactive'),
                    ('tbt', 'total-blocking-time'),
                    ('cls', 'cumulative-layout-shift'),
                    ('si', 'speed-index'),
                    ('ttfb', 'server-response-time'),
                    ('fmp', 'first-meaningful-paint')
                ]
                
                for metric_key, audit_key in metrics_to_check:
                    if audit_key in audits:
                        metrics[metric_key] = audits[audit_key]['displayValue']
                    else:
                        metrics[metric_key] = None  # Alterado para None

                metrics_list.append(metrics)

        except requests.exceptions.RequestException as e:
            print('Erro na requisição:', e)

    return metrics_list

# Função para ler URLs de um arquivo CSV
def read_urls_from_csv(file_path):
    urls = []
    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            urls.append(row[0])  # Supondo que as URLs estejam na primeira coluna
    return urls


# Exemplo de uso
csv_file = 'urls.csv'  # Substitua pelo caminho do seu arquivo CSV
urls = read_urls_from_csv(csv_file)
metrics = get_pagespeed_metrics(urls)

# Criação do arquivo Excel
wb = Workbook()
ws = wb.active

# Cabeçalho
header = ['url', 'score', 'fcp', 'lcp', 'tti', 'tbt', 'cls', 'si', 'ttfb', 'fmp', 'data']
ws.append(header)

# Exibindo as métricas
for metric in metrics:
    print('URL:', metric['url'])
    for key, value in metric.items():
        if key != 'url':
            print(key.upper() + ':', value if value is not None else 'N/A')  # Exibe 'N/A' se o valor for None
    print()
    
# Adiciona os dados
current_date = datetime.datetime.now().strftime('%Y-%m-%d')
for metric in tqdm(metrics, desc='Exportando', unit='Métrica'):
    metric['data'] = current_date
    row_data = [metric[field] for field in header]
    ws.append(row_data)

# Salva o arquivo
xlsx_export = f'relatorio_PSI_{current_date}.xlsx'
wb.save(xlsx_export)

print(f'Relatório exportado para {xlsx_export}')

input("Pressione 'Enter' para fechar o programa.")
