import docx2pdf
import os

# Solicita ao usuário que informe a pasta de origem
origem = input("Informe o caminho da pasta de origem: ")

# Solicita ao usuário que informe a pasta de destino
destino = input("Informe o caminho da pasta de destino: ")

# Muda para a pasta de origem
os.chdir(origem)

# Obtém a lista de todos os arquivos docx da pasta de origem
arquivos_docx = [arquivo for arquivo in os.listdir() if arquivo.endswith(".docx")]

# Percorre cada arquivo docx e o converte para PDF
for arquivo in arquivos_docx:
    docx2pdf.convert(arquivo, f"{destino}/{arquivo.replace('.docx', '.pdf')}")
