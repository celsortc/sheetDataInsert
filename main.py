import pdfplumber
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
import os

with pdfplumber.open("TCE - Izabela Sophia da Silva.pdf") as pdf:
    texto = ""

    for pagina in pdf.pages:
        texto += pagina.extract_text() + "\n"


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

    dados["horario-entrada"] = entrada
    dados["horario-saida"] = saida

    entrada_dt = datetime.strptime(entrada, "%H:%M")
    saida_dt = datetime.strptime(saida, "%H:%M")

    carga = (saida_dt - entrada_dt).total_seconds() / 3600
    dados["carga-horaria"] = f"{carga:.0f}"

fones = re.findall(r"Fone:\s*([^\n]+)", texto)

if len(fones) >= 3:
    dados["telefone"] = fones[2].strip()
else:
    dados["telefone"] = None

ordem = [
    "data-inicial",
    "data-final",
    "vazio",
    "nome",
    "CPF",
    "ano/curso",
    "horario-entrada",
    "horario-saida",
    "carga-horaria",
    "supervisor",  # agora aparece depois da carga horaria
    "telefone"
]


arquivo = "estags.xlsx"

if os.path.exists(arquivo):
    wb = load_workbook(arquivo)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.title = "Estags"
    ws.append(ordem)

linha = [dados.get(campo) for campo in ordem]
ws.append(linha)

wb.save(arquivo)

for campo in ordem:
    print(campo, "->", dados.get(campo))

