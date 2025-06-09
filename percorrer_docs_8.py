import fitz  # PyMuPDF
import re
import os
import glob


# Caminho da pasta com os arquivos PDF
pasta = "C:\\Users\\User\\OneDrive\\Documentos\\Programação\\Python\\Nova pasta\\documentos"

pattern_corrigido = re.compile(
    r'\n?\s*\d+\s+((?:[A-ZÁÉÍÓÚÂÊÔÃÕÇa-z]+\s*){2,10})\s+Nota:\s*(.*?)\s+RP:',
    re.DOTALL | re.IGNORECASE
)


arquivos_pdf = glob.glob(os.path.join(pasta, '*.pdf'))

for caminho_pdf in arquivos_pdf:
    pdf_document = fitz.open(caminho_pdf)
    full_text = ""
    for page in pdf_document:
        full_text += page.get_text("text")
    pdf_document.close()
    print(full_text)

    full_text = re.sub(r'[\t\r\f\v]+', ' ', full_text)
    full_text = re.sub(r' {2,}', ' ', full_text)

   
    results = []

    matches = pattern_corrigido.findall(full_text)

    print(matches)
  
    for nome_raw, notas_raw in matches:
        print(notas_raw,"\n", nome_raw)
        nome = ' '.join(nome_raw.split())
        linhas = [linha.strip() for linha in notas_raw.splitlines() if linha.strip()]
        notas = re.findall(r'\d+,\d+|\d+\.\d+|\d+', notas_raw)  
        mb = notas[-3].replace(',', '.') if notas else '0'
        results.append(f"{nome}    {mb}")


    nome_arquivo_pdf = os.path.splitext(os.path.basename(caminho_pdf))[0]
    caminho_saida = os.path.join(pasta, f"{nome_arquivo_pdf}_resultado.txt")

    with open(caminho_saida, 'w', encoding='utf-8') as f:
        for res in results:
            f.write(res + '\n')



print("Todos os arquivos PDF foram processados com sucesso!")
