# Importando as bibliotecas necessárias
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
from openpyxl import Workbook
import re

def main(planilha):
    """
    Extrai informações de um PDF e as armazena em um arquivo Excel.

    :param planilha: Nome do arquivo Excel de saída.
    :type planilha: str
    """

    texto = []
    for page_layout in extract_pages('pdfs\pdf-1.pdf'):
        for element in page_layout:
            if isinstance(element,LTTextContainer):
                dados = element.get_text().split('\n')
                # Organização é extraída se não houver cabeçalho de página
                organizacao = dados[0] if len(dados) < 3 and dados[0] != ' ' else organizacao
                # Processamento de informações da instituição
                if len(dados) > 2:
                    instituicao,cargo,responsavel,telefone,email,endereco = None,None,None,None,None,None
                    for dado in dados[:-1]:
                        instituicao = dado if not re.findall(r':',dado) else instituicao
                        cargo = dado.split(':')[0] if not dado == instituicao and not re.findall(r'Telefone:',dado) and not re.findall(r'Email',dado) and not re.findall(r'Endereço:',dado) else cargo
                        responsavel = dado.split(':')[1] if not dado == instituicao and not re.findall(r'Telefone:',dado) and not re.findall(r'Email',dado) and not re.findall(r'Endereço:',dado) else responsavel
                        telefone = dado.replace('Telefone: ','').replace('Celular: ','') if re.findall(r'Telefone: ',dado) else telefone
                        email = dado.replace('Email institucional: ','').replace('Email: ','') if re.findall(r'Email',dado) else email
                        endereco = dado.replace('Endereço: ','') if re.findall(r'Endereço: ',dado) else endereco
                    texto.append([organizacao,instituicao,cargo,responsavel,telefone,email,endereco])

    # Criação de um arquivo Excel e uma planilha para armazenar os dados
    workbook = Workbook()
    workbook['Sheet'].title = 'Dados PDF'
    sheet = workbook['Dados PDF']
    sheet.append(['Organizacao','Instituicao','Cargo','Resposavel','Telefone','Email','Endereco'])

    # Preenchimento da planilha com os dados extraídos
    for linha in texto:
        sheet.append(linha)
    # Salvamento do arquivo Excel com o nome fornecido
    workbook.save(f'tabelas\{planilha}.xlsx')

if __name__ == '__main__':
    main('dados-PDF-1')