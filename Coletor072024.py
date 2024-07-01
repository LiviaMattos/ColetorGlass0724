import tkinter as tk
from tkinter import messagebox, simpledialog
import os
import time
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
import logging
from dateutil import parser

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Definindo URLs globais
url_base = "https://www.glassdoor.com.br/Avalia%C3%A7%C3%B5es/BRQ-Avalia%C3%A7%C3%B5es-E224434.htm?sort.sortType=RD&sort.ascending=false&filter.iso3Language=por"
url_template = "https://www.glassdoor.com.br/Avalia%C3%A7%C3%B5es/BRQ-Avalia%C3%A7%C3%B5es-E224434_P{}.htm?sort.sortType=RD&sort.ascending=false&filter.iso3Language=por"

# Mapeamento de meses em português para números
meses = {
    'jan.': '01',
    'fev.': '02',
    'mar.': '03',
    'abr.': '04',
    'mai.': '05',
    'jun.': '06',
    'jul.': '07',
    'ago.': '08',
    'set.': '09',
    'out.': '10',
    'nov.': '11',
    'dez.': '12'
}

# Função para substituir os meses em português pelo seu número correspondente
def substituir_mes(data_str):
    for mes, num in meses.items():
        if mes in data_str:
            return data_str.replace(mes, num)
    return data_str

# Função para obter avaliações do site
def obter_avaliacoes_selenium(url, driver):
    try:
        driver.get(url)
        logging.info("Página carregada com sucesso!")
        time.sleep(5)
        conteudo = driver.page_source
        logging.info("Conteúdo da página obtido com sucesso")

        soup = BeautifulSoup(conteudo, 'html.parser')
        avaliacoes = []

        class Avaliacao:
            def __init__(self, data, titulo, nota, cargo, pros, contras):
                self.data = data
                self.titulo = titulo
                self.nota = nota
                self.cargo = cargo
                self.pros = pros
                self.contras = contras

        for avaliacao in soup.find_all('div', class_='review-details_topReview__5NRVX'):
            try:
                titulo_avaliacao = avaliacao.find('div', class_='review-details_titleHeadline__Jppto').get_text(strip=True)
                nota_avaliacao = avaliacao.find('span', class_='review-details_overallRating__Rxhdr').get_text(strip=True)
                cargo_avaliacao = avaliacao.find('span', class_='review-details_employee__MeSp3').get_text(strip=True)
                data_avaliacao = avaliacao.find('span', class_='timestamp_reviewDate__fBGY6').get_text(strip=True)
                pros_avaliacao = avaliacao.find('span', attrs={'data-test': 'review-text-pros'}).get_text(strip=True)
                contras_avaliacao = avaliacao.find('span', attrs={'data-test': 'review-text-cons'}).get_text(strip=True)

                avaliacoes.append(
                    Avaliacao(data_avaliacao, titulo_avaliacao, nota_avaliacao, cargo_avaliacao, pros_avaliacao, contras_avaliacao))
            except Exception as e:
                logging.error(f"Erro ao processar uma avaliação: {e}")

        return avaliacoes
    except Exception as e:
        logging.error(f"Erro ao obter avaliações do site: {e}")
        raise

def salvar_em_planilha(avaliacoes, nome_arquivo=None):
    try:
        data = {
            'Data': [avaliacao.data for avaliacao in avaliacoes],
            'Título': [avaliacao.titulo for avaliacao in avaliacoes],
            'Nota': [avaliacao.nota for avaliacao in avaliacoes],
            'Cargo': [avaliacao.cargo for avaliacao in avaliacoes],
            'Pros': [avaliacao.pros for avaliacao in avaliacoes],
            'Contras': [avaliacao.contras for avaliacao in avaliacoes]
        }

        df = pd.DataFrame(data)

        # Imprimir a coluna de datas original para depuração
        print("Datas originais:")
        print(df['Data'])

        # Substituir meses em português antes da conversão
        df['Data'] = df['Data'].apply(substituir_mes)

        # Tentativa de converter as datas para o formato DD/MM/AAAA
        def parse_date(date_str):
            try:
                return datetime.strptime(date_str, '%d de %m de %Y').strftime('%d/%m/%Y')
            except Exception as e:
                print(f"Erro ao converter data: {e}")
                return None

        df['Data'] = df['Data'].apply(parse_date)

        # Imprimir a coluna de datas após a tentativa de conversão
        print("Datas convertidas:")
        print(df['Data'])

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"avaliacoes_glassdoor_{timestamp}.xlsx"

        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
        caminho_arquivo = os.path.join(diretorio_atual, nome_arquivo)

        if os.path.exists(caminho_arquivo):
            dados_existentes = pd.read_excel(caminho_arquivo)
            dados_atualizados = pd.concat([dados_existentes, df], ignore_index=True)
        else:
            dados_atualizados = df

        dados_atualizados.to_excel(caminho_arquivo, index=False)
        logging.info(f"Arquivo salvo como {caminho_arquivo}")
    except Exception as e:
        logging.error(f"Erro ao salvar planilha: {e}")
        raise

def coletar_historico():
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    avaliacoes_totais = []

    try:
        avaliacoes_base = obter_avaliacoes_selenium(url_base, driver)
        avaliacoes_totais.extend(avaliacoes_base)

        for pagina in range(2, 10):
            url_atual = url_template.format(pagina)
            avaliacoes_pagina = obter_avaliacoes_selenium(url_atual, driver)
            if not avaliacoes_pagina:
                logging.info(f"Nenhuma avaliação encontrada na página {pagina}")
                break
            avaliacoes_totais.extend(avaliacoes_pagina)
            logging.info(f"Avaliações até a página {pagina} coletadas.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao coletar das páginas: {str(e)}")
    finally:
        driver.quit()
        if avaliacoes_totais:
            salvar_em_planilha(avaliacoes_totais)
            messagebox.showinfo("Sucesso", "Histórico de avaliações coletadas e salvas com sucesso!")
        else:
            messagebox.showinfo("Informação", "Nenhuma avaliação nova foi encontrada.")

def iniciar_coleta():
    url = url_input.get()
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        driver = webdriver.Chrome(options=options)
        avaliacoes_obtidas = obter_avaliacoes_selenium(url, driver)
        salvar_em_planilha(avaliacoes_obtidas)
        messagebox.showinfo("Sucesso", "Avaliações coletadas e salvas com sucesso em " + os.path.dirname(os.path.abspath(__file__)))
    except Exception as e:
        messagebox.showerror("Erro", str(e))

def mudar_url():
    global url_base
    url_base = simpledialog.askstring("Mudar URL", "Digite a nova URL:", initialvalue=url_base)
    url_input.delete(0, tk.END)
    url_input.insert(0, url_base)

if __name__ == '__main__':
    root = tk.Tk()
    root.title("Coletor de Avaliações")

    url_input = tk.Entry(root, width=100)
    url_input.insert(0, url_base)
    url_input.pack()

    botao_coletar = tk.Button(root, text="Coletar Avaliações", command=iniciar_coleta)
    botao_coletar.pack()

    botao_coletar_historico = tk.Button(root, text="Coletar Histórico", command=coletar_historico)
    botao_coletar_historico.pack()

    botao_mudar_url = tk.Button(root, text="Mudar URL", command=mudar_url)
    botao_mudar_url.pack()

    root.mainloop()

