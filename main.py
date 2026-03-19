import pdfplumber
import re

with pdfplumber.open("TCE - Izabela Sophia da Silva.pdf") as pdf:
    texto = ""

    for pagina in pdf.pages:
        texto += pagina.extract_text() + "\n"

dados = {
    "data-inicial": r"Vigência de:\s*(.*?)\sAté",
    "data-final": r"até\s*(.*)",
    "nome": r"Nome:\s*(.*?)\s+Código",
    "nome": r"Nome:\s*(.*?)\s+Código",

}

for campo, padrao in dados.items():
    match = re.search(padrao, texto, re.IGNORECASE)

    if match:
        print(campo, "->", match.group(1).strip())
    else:
        print(campo, "não encontrado")