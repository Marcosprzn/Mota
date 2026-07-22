import os
import sys
from xlwt import Workbook, easyxf
from collections import OrderedDict

# Adicionar a pasta dashboard ao sys.path para importar o processador
dashboard_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dashboard')
sys.path.append(dashboard_dir)

from processador import processar

print('Processando todas as pastas em dados/ ...')
dados = processar()
resultados_dicts = dados['funcionarios']

# Converter lista de dicionários para o formato de lista esperado pelo Excel
resultados = []
for d in resultados_dicts:
    resultados.append([
        str(d['id']), d['nome'], d['engenho'], d['turma'], d['media'],
        d['folgas'], d['faltas'], d['perder'], d['garantidos'],
        d['tipo'], d['producao'], d['valor_folgas'], d['total']
    ])

print(f'  Processados {len(resultados)} funcionários no total (de todos os períodos)')

# ── 3. Gerar arquivo de saída formatado ──
print('Gerando arquivo formatado...')
arquivo_saida = 'resultado_folgas.xls'

# ── Agrupar por engenho (preservando ordem) ──
engenhos = OrderedDict()
for row in resultados:
    eng = row[2]
    if eng not in engenhos:
        engenhos[eng] = []
    engenhos[eng].append(row)

# ── Estilos ──
base_fmt = 'font: name Calibri, height 200; alignment: vert centre; borders: left thin, right thin, top thin, bottom thin;'

style_cab = easyxf(
    'font: name Calibri, height 220, bold on, colour white;'
    'pattern: pattern solid, fore_colour dark_blue;'
    'alignment: horiz centre, vert centre, wrap on;'
    'borders: left thin, right thin, top thin, bottom thin;'
)

style_num = easyxf(base_fmt + 'alignment: horiz centre, vert centre;')
style_num.num_format_str = '#,##0.00'

style_int = easyxf(base_fmt + 'alignment: horiz centre, vert centre;')
style_int.num_format_str = '#,##0'

style_text = easyxf(base_fmt + 'alignment: horiz left, vert centre;')

alt_fmt = 'font: name Calibri, height 200; pattern: pattern solid, fore_colour light_green; alignment: vert centre; borders: left thin, right thin, top thin, bottom thin;'

style_num_alt = easyxf(alt_fmt + 'alignment: horiz centre, vert centre;')
style_num_alt.num_format_str = '#,##0.00'

style_int_alt = easyxf(alt_fmt + 'alignment: horiz centre, vert centre;')
style_int_alt.num_format_str = '#,##0'

style_text_alt = easyxf(alt_fmt + 'alignment: horiz left, vert centre;')

tot_fmt = 'font: name Calibri, height 220, bold on, colour white; pattern: pattern solid, fore_colour dark_blue; alignment: horiz centre, vert centre; borders: left medium, right medium, top medium, bottom medium;'

style_total = easyxf(tot_fmt)
style_total.num_format_str = '#,##0.00'

style_total_txt = easyxf(tot_fmt)

cabecalhos = [
    'ID', 'Nome', 'Engenho', 'Turma', 'Média',
    'Folgas/\nFeriados', 'Faltas', 'Repousos\nPerder',
    'Repousos\nGarantidos', 'Tipo',
    'Produção (R$)',
    'Valor Folgas/\nFeriados (R$)', 'Soma Total (R$)'
]

col_widths = [10, 40, 22, 8, 10, 10, 8, 10, 10, 10, 18, 18, 18]

def escrever_planilha(ws, dados, titulo=None):
    # Cabeçalho
    for c, (tit, w) in enumerate(zip(cabecalhos, col_widths)):
        ws.write(0, c, tit, style_cab)
        ws.col(c).width = int(w * 256)

    if not dados:
        return

    # Dados
    for i, row_data in enumerate(dados):
        is_alt = i % 2 == 1
        r = i + 1
        ws.write(r, 0, int(float(row_data[0])), style_int_alt if is_alt else style_int)
        ws.write(r, 1, row_data[1], style_text_alt if is_alt else style_text)
        ws.write(r, 2, row_data[2], style_text_alt if is_alt else style_text)
        ws.write(r, 3, row_data[3], style_text_alt if is_alt else style_text)
        ws.write(r, 4, row_data[4], style_num_alt if is_alt else style_num)
        ws.write(r, 5, row_data[5], style_int_alt if is_alt else style_int)   # Folgas/Feriados
        ws.write(r, 6, row_data[6], style_int_alt if is_alt else style_int)   # Faltas
        ws.write(r, 7, row_data[7], style_int_alt if is_alt else style_int)   # Repousos Perder
        ws.write(r, 8, row_data[8], style_int_alt if is_alt else style_int)   # Folgas Ajustadas
        ws.write(r, 9, row_data[9], style_text_alt if is_alt else style_text)  # Tipo
        ws.write(r, 10, row_data[10], style_num_alt if is_alt else style_num)  # Producao
        ws.write(r, 11, row_data[11], style_num_alt if is_alt else style_num)  # Valor Folgas
        ws.write(r, 12, row_data[12], style_num_alt if is_alt else style_num)  # Soma Total

    # Linha de total
    tr = len(dados) + 1
    ws.write(tr, 0, '', style_total_txt)
    ws.write(tr, 1, 'TOTAIS' + (f' - {titulo}' if titulo else ''), style_total_txt)
    for c in range(2, 5):
        ws.write(tr, c, '', style_total_txt)
    ws.write(tr, 5, sum(r[5] for r in dados), style_total)   # Folgas
    ws.write(tr, 6, sum(r[6] for r in dados), style_total)   # Faltas
    ws.write(tr, 7, sum(r[7] for r in dados), style_total)   # Repousos Perder
    ws.write(tr, 8, sum(r[8] for r in dados), style_total)   # Repousos Garantidos
    ws.write(tr, 9, '', style_total_txt)                      # Tipo
    ws.write(tr, 10, sum(r[10] for r in dados), style_total) # Producao
    ws.write(tr, 11, sum(r[11] for r in dados), style_total) # Valor Folgas
    ws.write(tr, 12, sum(r[12] for r in dados), style_total) # Soma Total

wb_out = Workbook(encoding='utf-8')

# ── Aba GERAL ──
ws_geral = wb_out.add_sheet('Geral')
escrever_planilha(ws_geral, resultados, 'GERAL')

# ── Uma aba para cada ENGENHO ──
import re
def sanitizar_aba(nome):
    nome = re.sub(r'[\\/?*\[\]:]', '', nome)
    return nome[:31] or 'Aba'

for eng, dados_eng in engenhos.items():
    nome_aba = sanitizar_aba(eng)
    ws_eng = wb_out.add_sheet(nome_aba)
    escrever_planilha(ws_eng, dados_eng, eng)

wb_out.save(arquivo_saida)
print(f'Arquivo salvo: {arquivo_saida} com {1 + len(engenhos)} abas')

# ── 4. Exibir resumo ──
print('\n———  RESUMO  ———')
print(f'{"ID":>8} {"Nome":<35} {"Eng":<12} {"Trm":>4} {"Media":>8} {"Folgas":>6} {"Faltas":>6} {"Perder":>6} {"Garant":>6} {"Tipo":>8} {"Producao":>10} {"V.Folgas":>10} {"Total":>10}')
print('-' * 145)
for row in resultados[:5]:
    print(f'{row[0]:>8} {row[1]:<35} {row[2]:<12} {row[3]:>4} {row[4]:>8.2f} {row[5]:>6} {row[6]:>6} {row[7]:>6} {row[8]:>6} {row[9]:>8} {row[10]:>10.2f} {row[11]:>10.2f} {row[12]:>10.2f}')
if len(resultados) > 5:
    print(f'... e mais {len(resultados) - 5} funcionários')
print(f'\nTotal de funcionários: {len(resultados)}')
print(f'Total Folgas/Feriados: {sum(r[5] for r in resultados)}')
print(f'Total Faltas: {sum(r[6] for r in resultados)}')
print(f'Total Repousos Perder: {sum(r[7] for r in resultados)}')
print(f'Total Repousos Garantidos: {sum(r[8] for r in resultados)}')
meias = sum(1 for r in resultados if r[9] == 'MEIA')
print(f'Funcionários MEIA período: {meias}')
print(f'Soma Produção: R$ {sum(r[10] for r in resultados):.2f}')
print(f'Soma Valor Folgas: R$ {sum(r[11] for r in resultados):.2f}')
print(f'Soma Total Geral: R$ {sum(r[12] for r in resultados):.2f}')
