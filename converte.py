import os
import win32com.client

# especificar a pasta que contém os arquivos do Word
folder = r'C:\Users\jepim\Desktop\testeGPT\peças corrigidas'

# especificar a pasta para salvar os arquivos PDF
pdf_folder = r'C:\Users\jepim\Desktop\testeGPT\para peticionar'

# inicializar o Microsoft Word
word = win32com.client.Dispatch('Word.Application')

# percorrer todos os arquivos na pasta especificada
for filename in os.listdir(folder):
    # verificar se o arquivo é do Word
    if filename.endswith('.docx'):
        # criar o caminho completo do arquivo
        filepath = os.path.join(folder, filename)
        # abrir o arquivo no Word
        doc = word.Documents.Open(filepath)
        # gerar o nome do arquivo pdf
        pdf_name = os.path.splitext(filename)[0] + '.pdf'
        # gerar o caminho completo do arquivo pdf
        pdf_path = os.path.join(pdf_folder, pdf_name)
        # salvar o arquivo como pdf
        doc.SaveAs(pdf_path, FileFormat=17)
        # fechar o arquivo
        doc.Close()

# fechar o Microsoft Word
word.Quit()
