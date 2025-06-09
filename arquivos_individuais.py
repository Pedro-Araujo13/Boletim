import os
import glob

class Aluno:
    def __init__(self, nome_original):
        self.nome_original = nome_original  # Nome como aparece nos arquivos
        self.notas = []

    def adicionar_nota(self, nota):
        self.notas.append(nota)

    def __str__(self):
        notas_str = '    '.join(f"{nota:.1f}" for nota in self.notas)
        return f"{self.nome_original}    {notas_str}"

def processar_arquivo(nome_arquivo, dicionario_alunos):
    with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
        for linha in arquivo:
            linha = linha.strip()
            partes = linha.rsplit(' ', 1)
            if len(partes) == 2:
                nome, nota_str = partes
                nota = float(nota_str.replace(',', '.'))
                
                # Limpeza: remove espaços e coloca tudo em minúsculo para chave
                chave_nome = nome.strip().lower()
                
                if chave_nome in dicionario_alunos:
                    dicionario_alunos[chave_nome].adicionar_nota(nota)
                else:
                    aluno = Aluno(nome.strip())
                    aluno.adicionar_nota(nota)
                    dicionario_alunos[chave_nome] = aluno

# Caminho da pasta com os arquivos
pasta = "C:\\Users\\User\\OneDrive\\Documentos\\Programação\\Python\\ConversorPDFtxt\\arquivostxt"  # <-- MODIFIQUE AQUI

# Encontra todos os arquivos .txt na pasta
arquivos_txt = glob.glob(os.path.join(pasta, '*.txt'))

# Dicionário para armazenar alunos
alunos = {}

# Processa todos os arquivos encontrados
for nome_arquivo in arquivos_txt:
    processar_arquivo(nome_arquivo, alunos)

# Para cada aluno, cria um arquivo individual
for chave_nome, aluno in alunos.items():
    # Nome do arquivo: "Pedro Lima de Araujo.txt"
    nome_arquivo = os.path.join(pasta, f"{aluno.nome_original}.txt")
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        f.write(str(aluno))

print("Arquivos individuais criados com sucesso!")