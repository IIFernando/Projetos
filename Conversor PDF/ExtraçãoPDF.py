import glob
import PyPDF2
import re

caminho = input(r'InfoG:\Meu Drive\Projetos\Conversor PDF\ExtratoPDF.pyrme a pasta dos arquivos: ')

arquivos = glob.glob(caminho + "/*.pdf")

# Abre o arquivo CSV fora do loop
with open('ExtraçãoPDF.csv', 'a', encoding='UTF-16') as arquivo:
    for tabela in arquivos:
        df = PyPDF2.PdfReader(tabela).pages[0].extract_text()
        df = re.sub('\n', '|', df)

        # Concatena o texto e o caminho do arquivo com o método join
        linha = '|'.join([df, tabela])

        # Grava a linha no arquivo CSV
        arquivo.write(linha + '\n')

# Fecha o arquivo CSV após o loop
arquivo.close()