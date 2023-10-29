# Importando as bibliotecas necessárias
from pdfminer.high_level import extract_text, extract_pages
from pdfminer.layout import LTTextContainer
from openpyxl import Workbook
import re

def main(planilha):
    """
    Extrai informações de um PDF e as armazena em um arquivo Excel.

    :param planilha: Nome do arquivo Excel de saída.
    :type planilha: str
    """

    matrix_dados = []
    for page_layout in extract_pages('pdfs\pdf-2.pdf'):
        for element in page_layout:
            if isinstance(element,LTTextContainer):
                texto = re.sub(r'(?<=\S)\s(?=\S)','',element.get_text())
                texto_ = re.sub(r'\s+',' ',texto)
                matrix_dados.append(texto_.split('\n'))

    # Filtra os dados da matriz
    matrix_filtrada = [dado for dado in matrix_dados if not dado[0] == ' '][2:]

    # Identifica as instituições e suas informações
    instituicao = [[id,dado[0]] for id,dado in enumerate(matrix_filtrada) if not re.findall(r'(TELEFONE:)',dado[0]) and not re.findall(r'(http)',dado[0]) and not re.findall(r'(E-mail:)',dado[0]) and not re.findall(r'(Saiba)',dado[0])]

    # Extrai os dados das instituições
    dados = [matrix_filtrada[linha[0]+1:instituicao[id+1][0]] if id+1 < len(instituicao) else matrix_filtrada[linha[0]+1:] for id,linha in enumerate(instituicao)]

    # Cria um arquivo Excel e uma planilha para armazenar os dados
    workbook = Workbook()
    workbook['Sheet'].title = 'Dados-PDF'
    sheet = workbook['Dados-PDF']
    sheet.append(['Instituição','Telefone','E-mail','Site'])

    # Preenche a planilha com os dados
    for id,linha in enumerate(dados):
        telefone = ''.join(item[0].replace('TELEFONE: ','') for item in linha if re.findall(r'(TELEFONE:)', item[0]))
        email = ''.join(re.sub(r'\s+', '', item[0].replace('E-mail: ','')) for item in linha if re.findall(r'(E-mail)', item[0]))
        link = ''.join(re.sub(r'\s+', '', item[0].replace('Saiba mais:','').replace('FALE CONOSCO ON-LINE:','').replace('Saibamais:','')) for item in linha if re.findall(r'(http)', item[0]))
        sheet.append([instituicao[id][1],telefone,email,link])

    # Salva o arquivo Excel
    workbook.save(f'tabelas\{planilha}.xlsx')

# Chama a função principal com o nome do arquivo de saída desejado
if __name__ == '__main__':
    main('dados-PDF-2')