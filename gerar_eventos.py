import pandas as pd
from datetime import datetime, timedelta
import os

# Nome do arquivo da planilha
planilha = "certificados_tsplus.xlsx"
saida_pasta = "eventos_ics/"
pasta_recentes = "eventos_recentes_ics/"

# Nome dos arquivos consolidados
arquivo_todos_eventos = "todos_eventos.ics"
arquivo_recentes = "renovados_recentes.ics"

# Garantir que as pastas existam
os.makedirs(saida_pasta, exist_ok=True)
os.makedirs(pasta_recentes, exist_ok=True)

# Ler a planilha
df = pd.read_excel(planilha)

# Carregar eventos existentes em `todos_eventos.ics`
eventos_existentes = {}
if os.path.exists(arquivo_todos_eventos):
    with open(arquivo_todos_eventos, "r") as f:
        conteudo = f.read()
        eventos = conteudo.split("BEGIN:VEVENT")
        for evento in eventos[1:]:
            linhas = evento.split("\n")
            nome_evento = next((linha.replace("SUMMARY:", "").strip() for linha in linhas if linha.startswith("SUMMARY:")), None)
            data_expiracao = next((linha.replace("DTSTART;VALUE=DATE:", "").strip() for linha in linhas if linha.startswith("DTSTART;VALUE=DATE:")), None)
            if nome_evento and data_expiracao:
                eventos_existentes[nome_evento] = data_expiracao

# Gerar eventos
eventos_recentes = []  # Para consolidar eventos novos ou alterados
for index, row in df.iterrows():
    nome_tsplus = row["Nome"]
    data_expiracao = pd.to_datetime(row["Data de Expiração"]).strftime('%Y%m%d')  # Converter para formato esperado
    titulo = f"[NOC] - Renovar certificado {nome_tsplus}"

    # Verificar se é novo ou se a data de expiração foi alterada
    if titulo not in eventos_existentes or eventos_existentes[titulo] != data_expiracao:
        # Criar o conteúdo do arquivo .ics
        evento = f"""BEGIN:VEVENT
SUMMARY:{titulo}
DESCRIPTION:Renovar certificado TSPLUS
DTSTART;VALUE=DATE:{data_expiracao}
DTEND;VALUE=DATE:{data_expiracao}
DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}
BEGIN:VALARM
TRIGGER:-P1D
ACTION:DISPLAY
DESCRIPTION:Lembrete: {titulo}
END:VALARM
ATTENDEE;CN="joao.manduca@dorpa.com.br":mailto:joao.manduca@dorpa.com.br
END:VEVENT
"""

        # Nome do arquivo .ics individual
        nome_arquivo_recente = f"{pasta_recentes}evento_{nome_tsplus}.ics"

        # Salvar o arquivo individual na pasta de recentes
        with open(nome_arquivo_recente, "w") as file_recente:
            file_recente.write(f"BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n{evento}\nEND:VCALENDAR")
        print(f"Evento recente salvo: {nome_arquivo_recente}")

        # Adicionar ao consolidado recente
        eventos_recentes.append(evento)

# Consolidar os eventos recentes em `renovados_recentes.ics`
with open(arquivo_recentes, "w") as f_recentes:
    f_recentes.write("BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n")
    for evento in eventos_recentes:
        f_recentes.write(evento)
    f_recentes.write("END:VCALENDAR")

print(f"Arquivo consolidado de eventos recentes criado: {arquivo_recentes}")

# Atualizar o arquivo `todos_eventos.ics` com os eventos recentes
with open(arquivo_todos_eventos, "a") as f_todos:
    for evento in eventos_recentes:
        f_todos.write(evento)

print(f"Arquivo consolidado atualizado: {arquivo_todos_eventos}")
