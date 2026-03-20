import pdfplumber
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
import os


ordem = [
    "data-inicial",
    "data-final",
    "vazio",
    "nome",
    "CPF",
    "ano/curso",
    "hr-entrada-saida",
    "carga-horaria",
    "supervisor",  # agora aparece depois da carga horaria
    "telefone"
]

# Define a pasta que será aberta e lista os pdfs
pasta = "pdfs" 

arquivos = [f for f in os.listdir(pasta) if f.endswith(".pdf")]

arquivo = "estags.xlsx"

if os.path.exists(arquivo):
    wb = load_workbook(arquivo)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.title = "Estags"
    ws.append(ordem)

for arquivo in arquivos:
    print("Processando:", arquivo)

    caminho_pdf = os.path.join(pasta, arquivo)

    with pdfplumber.open(caminho_pdf) as pdf:
        texto = ""
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"




    # função que vai formatar o nome, permitir que capitalize o nome com excecão de palavras como de, do, dos..
    def formatar_nome(nome):
        if not nome:
            return nome

        palavras_minusculas = {"da", "de", "do", "das", "dos"}

        palavras = nome.lower().split()
        resultado = []

        for i, palavra in enumerate(palavras):
            if palavra in palavras_minusculas and i != 0:
                resultado.append(palavra)
            else:
                resultado.append(palavra.capitalize())

        return " ".join(resultado)

    def formatar_numeros(cpf):
        if not cpf:
            return cpf
        
        return re.sub(r'\D', '', cpf)

    def pegar_numero(nome_arquivo):
        numero = nome_arquivo.split(' - ')[0]
        return numero



    padroes = {
        "data-inicial": r"Vigência de:\s*(.*?)\sAté",
        "data-final": r"até\s*(.*)",
        "nome": r"Nome:\s*(.*?)\s+Código",
        "CPF": r"CPF/MF:\s*(.*)",
        "ano/curso": r"Regularmente Matriculado:\s*(\d+)",
        "supervisor": r"Supervisor:\s*(.*?)\sCargo",
        
    }

    dados = {}

    for campo, padrao in padroes.items():
        match = re.search(padrao, texto, re.IGNORECASE)

        if match:
            dados[campo] = match.group(1).strip()
        else:
            dados[campo] = None



    match_horario = re.search(
        r"Horário das\s*(\d{2}:\d{2})\s*as\s*(\d{2}:\d{2})",
        texto,
        re.IGNORECASE
    )

    if match_horario:
        entrada = match_horario.group(1)
        saida = match_horario.group(2)

        # dados["horario-entrada"] = entrada
        # dados["horario-saida"] = saida
        dados["hr-entrada-saida"] = entrada + " - " + saida
        print(entrada + "-"+saida)

        entrada_dt = datetime.strptime(entrada, "%H:%M")
        saida_dt = datetime.strptime(saida, "%H:%M")

        carga = (saida_dt - entrada_dt).total_seconds() / 3600
        dados["carga-horaria"] = f"{carga:.0f}"

    fones = re.findall(r"Fone:\s*([^\n]+)", texto)

    if len(fones) >= 3:
        dados["telefone"] = fones[2].strip()
    else:
        dados["telefone"] = None



    if dados.get("nome"):
        dados["nome"] = formatar_nome(dados["nome"])
    if dados.get("CPF"):
        dados["CPF"] = formatar_numeros(dados["CPF"])
    if dados.get("telefone"):
        dados["telefone"] = formatar_numeros(dados["telefone"])



    linha = [dados.get(campo) for campo in ordem]
    ws.append(linha)


    for campo in ordem:
        print(campo, "->", dados.get(campo))

arquivo_excel = "estag.xlsx"
wb.save(arquivo_excel)




