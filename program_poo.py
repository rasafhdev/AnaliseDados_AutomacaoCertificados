import pandas as pd
import openpyxl as xl
from PIL import Image, ImageDraw, ImageColor, ImageFont
import os
from time import sleep

class ProcessadorDados:
    # contrutor
    def __init__(self, caminho_csv):
        self.caminho_csv = caminho_csv
        self.df = None
    
    # limpar tela
    def limpar_tela(self):
        sleep(4)
        if os.name == 'nt':
            os.system('cls')
        else:
            os.system('clear')
    
    # carregando os dados
    def carregar_dados(self):
        print('CARREGANDO DADOS...')
        self.limpar_tela()
        self.df = pd.read_csv('res/dados.csv')

    # amostras do dataset
    def amostras(self):
        print('CARREGANDO AMOSTRAS...')
        self.limpar_tela()
        print(f'Amostra -> 15 Registros \n {self.df.head(15)}\n')
        print(f'Contagem de dados ausêntes -> \n {self.df.isna().sum()}')
    
    # tratamento de dados
    def tratar_dados(self):
        print('INICIANDO TRATAMENTO DE DADOS...\n')
        sleep(3)

        # calcula estatisticas 
        data_emissao = self.df['data_emissao'].combine_first(self.df['data_fim'])
        moda_carga_horaria = self.df['carga_horaria'].mode()[0]
        media_aproveitamento = self.df['aproveitamento'].mean()

        # insere correção
        self.df['data_emissao'] = self.df['data_emissao'].fillna(data_emissao)
        self.df['carga_horaria'] = self.df['carga_horaria'].fillna(moda_carga_horaria).astype(int).astype(str)
        self.df['aproveitamento'] = self.df['aproveitamento'].fillna(media_aproveitamento).astype(int).astype(str)

        # corrige as datas para o formato DD/MM/YYYY para powerbi
        colunas = ['data_inicio', 'data_fim', 'data_emissao', 'data_nascimento'] #lista com os nomes das colunas
        for coluna in colunas:
            self.df[coluna] = pd.to_datetime(self.df[coluna], errors='coerce').dt.strftime('%d/%m/%Y')
        
        self.df.to_excel('res/dados.xlsx', index=False)

        print('\nAções Realizadas:\n1)Calculo de Estatísticas\n2)Substituição dos valores NAN\n3)Correção das colunas de data para MM/DD/AAAA')

class EmiteCertificados:
    def __init__(self, arquivo, pagina):
        self.arquivo = arquivo
        self.pagina = pagina
    
    def emitir_certificados(self):
        print('INICIANDO MODULO DE EMISSÃO DE CERTIFICADOS')
        # loop com desempacotamento
        for i, linha in enumerate(self.pagina.iter_rows(min_row=2)):
            nome_curso, nome_aluno, modalidade, data_inicio, data_fim, data_emissao, carga_horaria, aproveitamento, data_nascimento, *_ = linha

            # abre certificado
            certificado = Image.open('res/certificadomodelo.png')
            
            # cor especial para nome
            cor_do_nome = ImageColor.getrgb('#E4BF5A')

            # Fontes
            font_do_nome = ImageFont.truetype('res/PinyonScript-Regular.ttf', 110)
            font_geral = ImageFont.truetype('res/tahoma.ttf', 35)

            # ação para desenhar no certificado
            insere_info = ImageDraw.Draw(certificado)
            
            # coordenadas para inserir info no certificado
            insere_info.text((780, 630), nome_aluno.value, fill=cor_do_nome, font=font_do_nome)
            insere_info.text((964, 758), nome_curso.value, fill='White', font=font_geral)
            insere_info.text((750, 813), modalidade.value, fill='White', font=font_geral)
            insere_info.text((900, 868), aproveitamento.value, fill='White', font=font_geral)
            insere_info.text((519, 950), carga_horaria.value, fill='White', font=font_geral)
            insere_info.text((730, 950), data_inicio.value, fill='White', font=font_geral)
            insere_info.text((1032, 950), data_fim.value, fill='White', font=font_geral)
            insere_info.text((1320, 950), data_emissao.value, fill='White', font=font_geral)

            # salva certificado
            certificado.save(f'certificados/{i+1}_{nome_aluno.value}_certficado.png')

            # Bandeira
            print(f'Certificado {i+1} Criado!')

            


if __name__ == '__main__':

    processador = ProcessadorDados('res/dados.csv')
    processador.carregar_dados()
    processador.amostras() # pré tratamento
    
    processador.limpar_tela()
    
    processador.tratar_dados()
    processador.amostras() # pós tratamento

    processador.limpar_tela()
    arquivo = xl.load_workbook('res/dados.xlsx')
    pagina = arquivo['Sheet1']

    emite_certificado = EmiteCertificados(arquivo, pagina)
    emite_certificado.emitir_certificados()