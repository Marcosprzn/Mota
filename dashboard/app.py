from flask import Flask, jsonify, render_template, request, Response
import os
import threading
import webbrowser
from processador import processar, encontrar_planilhas
from collections import OrderedDict
import csv, io

app = Flask(__name__)

dados_cache = None
cache_mtimes = {}

def carregar_dados():
    global dados_cache, cache_mtimes
    try:
        caminho_rel, caminho_lan = encontrar_planilhas()
        mtime_rel = os.path.getmtime(caminho_rel)
        mtime_lan = os.path.getmtime(caminho_lan)
    except Exception:
        mtime_rel = 0
        mtime_lan = 0
        
    current_mtimes = {'rel': mtime_rel, 'lan': mtime_lan}
    
    if dados_cache is None or cache_mtimes != current_mtimes:
        dados_cache = processar()
        cache_mtimes = current_mtimes
    return dados_cache

@app.route('/')
def index():
    return render_template('index.html')

def filtrar_e_agrupar_funcionarios(funcs, eng_filter, periodos_filter):
    if eng_filter and eng_filter != 'Todos':
        funcs = [d for d in funcs if d['engenho'] == eng_filter]
        
    if periodos_filter:
        p_list = periodos_filter.split(',')
        funcs = [d for d in funcs if d['competencia'] in p_list]
        
    # Agrupar por ID para somar meses
    agrupados = {}
    for d in funcs:
        fid = d['id']
        if fid not in agrupados:
            agrupados[fid] = d.copy()
            agrupados[fid]['periodos_contagem'] = 1
        else:
            ag = agrupados[fid]
            ag['folgas'] += d['folgas']
            ag['faltas'] += d['faltas']
            ag['perder'] += d['perder']
            ag['garantidos'] += d['garantidos']
            ag['producao'] += d['producao']
            ag['valor_folgas'] += d['valor_folgas']
            ag['total'] += d['total']
            ag['media'] = round((ag['media'] * ag['periodos_contagem'] + d['media']) / (ag['periodos_contagem'] + 1), 2)
            ag['periodos_contagem'] += 1
            if d['tipo'] == 'MEIA':
                ag['tipo'] = 'MEIA' # Se teve meia em algum mes, marca como meia
                
    return list(agrupados.values())

@app.route('/api/resumo')
def api_resumo():
    dados = carregar_dados()
    funcs_raw = dados.get('funcionarios', [])
    eng_filter = request.args.get('engenho')
    per_filter = request.args.get('periodo')
    
    funcs = filtrar_e_agrupar_funcionarios(funcs_raw, eng_filter, per_filter)
    
    engenhos = OrderedDict()
    for d in funcs:
        eng = d['engenho']
        if eng not in engenhos:
            engenhos[eng] = {'count': 0, 'producao': 0, 'valor_folgas': 0, 'total': 0, 'garantidos': 0, 'faltas': 0, 'folgas': 0, 'meias': 0}
        engenhos[eng]['count'] += 1
        engenhos[eng]['producao'] += d['producao']
        engenhos[eng]['valor_folgas'] += d['valor_folgas']
        engenhos[eng]['total'] += d['total']
        engenhos[eng]['garantidos'] += d['garantidos']
        engenhos[eng]['faltas'] += d['faltas']
        engenhos[eng]['folgas'] += d['folgas']
        if d['tipo'] == 'MEIA':
            engenhos[eng]['meias'] += 1

    totais = {
        'funcionarios': len(funcs),
        'producao': sum(d['producao'] for d in funcs),
        'valor_folgas': sum(d['valor_folgas'] for d in funcs),
        'total_geral': sum(d['total'] for d in funcs),
        'garantidos': sum(d['garantidos'] for d in funcs),
        'faltas': sum(d['faltas'] for d in funcs),
        'folgas': sum(d['folgas'] for d in funcs),
        'meias': sum(1 for d in funcs if d['tipo'] == 'MEIA'),
        'engenhos': len(engenhos),
        'periodos_disponiveis': dados.get('periodos_disponiveis', []),
        'competencia': per_filter.replace(',', ' e ') if per_filter else dados.get('competencia', 'Desconhecida')
    }

    eng_list = []
    for nome, vals in engenhos.items():
        vals['nome'] = nome
        eng_list.append(vals)

    return jsonify({'totais': totais, 'engenhos': eng_list})

@app.route('/api/funcionarios')
def api_funcionarios():
    dados = carregar_dados()
    funcs_raw = dados.get('funcionarios', [])
    eng = request.args.get('engenho')
    per = request.args.get('periodo')
    
    funcs = filtrar_e_agrupar_funcionarios(funcs_raw, eng, per)
    return jsonify(funcs)

@app.route('/api/exportar/excel')
def exportar_excel():
    dados = carregar_dados()
    funcs_raw = dados.get('funcionarios', [])
    eng = request.args.get('engenho')
    per = request.args.get('periodo')
    
    funcs = filtrar_e_agrupar_funcionarios(funcs_raw, eng, per)

    try:
        from xlwt import Workbook, easyxf
        import re
        
        def sanitizar_aba(nome):
            nome = re.sub(r'[\\/?*\[\]:]', '', nome)
            return nome[:31] or 'Aba'

        wb = Workbook(encoding='utf-8')
        
        # ── Estilos ──
        base_fmt = 'font: name Calibri, height 180; alignment: vert centre; borders: left thin, right thin, top thin, bottom thin;'
        
        style_cab = easyxf(
            'font: name Calibri, height 200, bold on, colour white;'
            'pattern: pattern solid, fore_colour dark_blue;'
            'alignment: horiz centre, vert centre, wrap on;'
            'borders: left thin, right thin, top thin, bottom thin;'
        )
        
        style_num = easyxf(base_fmt + 'alignment: horiz centre, vert centre;')
        style_num.num_format_str = '#,##0.00'
        
        style_int = easyxf(base_fmt + 'alignment: horiz centre, vert centre;')
        style_int.num_format_str = '#,##0'
        
        style_text = easyxf(base_fmt + 'alignment: horiz left, vert centre;')
        
        alt_fmt = 'font: name Calibri, height 180; pattern: pattern solid, fore_colour light_green; alignment: vert centre; borders: left thin, right thin, top thin, bottom thin;'
        
        style_num_alt = easyxf(alt_fmt + 'alignment: horiz centre, vert centre;')
        style_num_alt.num_format_str = '#,##0.00'
        
        style_int_alt = easyxf(alt_fmt + 'alignment: horiz centre, vert centre;')
        style_int_alt.num_format_str = '#,##0'
        
        style_text_alt = easyxf(alt_fmt + 'alignment: horiz left, vert centre;')
        
        tot_fmt = 'font: name Calibri, height 200, bold on, colour white; pattern: pattern solid, fore_colour dark_blue; alignment: horiz centre, vert centre; borders: left medium, right medium, top medium, bottom medium;'
        
        style_total = easyxf(tot_fmt)
        style_total.num_format_str = '#,##0.00'
        
        style_total_txt = easyxf(tot_fmt)

        cabecalhos = [
            'ID', 'Nome', 'Engenho', 'Turma', 'Média',
            'Folgas', 'Faltas', 'Perder',
            'Garantidos', 'Tipo',
            'Produção', 'Vl.Total Repouso', 'Total'
        ]
        col_widths = [10, 40, 22, 8, 10, 8, 8, 8, 10, 8, 16, 16, 16]

        def escrever_planilha(ws, dados_plan):
            # Cabeçalho
            for c, (tit, w) in enumerate(zip(cabecalhos, col_widths)):
                ws.write(0, c, tit, style_cab)
                ws.col(c).width = int(w * 256)
            
            # Dados
            for i, d in enumerate(dados_plan):
                is_alt = i % 2 == 1
                r = i + 1
                ws.write(r, 0, int(d['id']), style_int_alt if is_alt else style_int)
                ws.write(r, 1, d['nome'], style_text_alt if is_alt else style_text)
                ws.write(r, 2, d['engenho'], style_text_alt if is_alt else style_text)
                ws.write(r, 3, d['turma'], style_text_alt if is_alt else style_text)
                ws.write(r, 4, d['media'], style_num_alt if is_alt else style_num)
                ws.write(r, 5, d['folgas'], style_int_alt if is_alt else style_int)
                ws.write(r, 6, d['faltas'], style_int_alt if is_alt else style_int)
                ws.write(r, 7, d['perder'], style_int_alt if is_alt else style_int)
                ws.write(r, 8, d['garantidos'], style_int_alt if is_alt else style_int)
                ws.write(r, 9, d['tipo'], style_text_alt if is_alt else style_text)
                ws.write(r, 10, d['producao'], style_num_alt if is_alt else style_num)
                ws.write(r, 11, d['valor_folgas'], style_num_alt if is_alt else style_num)
                ws.write(r, 12, d['total'], style_num_alt if is_alt else style_num)
            
            # Totais
            tr = len(dados_plan) + 1
            ws.write(tr, 0, '', style_total_txt)
            ws.write(tr, 1, 'TOTAIS', style_total_txt)
            for c in range(2, 5):
                ws.write(tr, c, '', style_total_txt)
            ws.write(tr, 5, sum(d['folgas'] for d in dados_plan), style_total)
            ws.write(tr, 6, sum(d['faltas'] for d in dados_plan), style_total)
            ws.write(tr, 7, sum(d['perder'] for d in dados_plan), style_total)
            ws.write(tr, 8, sum(d['garantidos'] for d in dados_plan), style_total)
            ws.write(tr, 9, '', style_total_txt)
            ws.write(tr, 10, sum(d['producao'] for d in dados_plan), style_total)
            ws.write(tr, 11, sum(d['valor_folgas'] for d in dados_plan), style_total)
            ws.write(tr, 12, sum(d['total'] for d in dados_plan), style_total)

        # Se for "Todos" e tiver múltiplos engenhos, gerar abas individuais além da aba "Geral"
        if not eng or eng == 'Todos':
            ws_geral = wb.add_sheet('Geral')
            escrever_planilha(ws_geral, funcs)
            
            # Agrupar por engenho
            engenhos = OrderedDict()
            for d in funcs:
                e = d['engenho']
                if e not in engenhos:
                    engenhos[e] = []
                engenhos[e].append(d)
                
            for e, dados_eng in engenhos.items():
                nome_aba = sanitizar_aba(e)
                ws_eng = wb.add_sheet(nome_aba)
                escrever_planilha(ws_eng, dados_eng)
        else:
            # Apenas uma aba com o engenho selecionado
            ws_eng = wb.add_sheet(sanitizar_aba(eng))
            escrever_planilha(ws_eng, funcs)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return Response(buf.getvalue(), mimetype='application/vnd.ms-excel',
            headers={'Content-Disposition': 'attachment; filename=dashboard_folgas.xls'})
    except Exception as e:
        print(f"Erro ao gerar planilha Excel (usando fallback CSV): {e}")

    # fallback CSV
    si = io.StringIO()
    w = csv.writer(si)
    w.writerow(['ID','Nome','Engenho','Turma','Vl.Repouso','Folgas','Faltas','Perder','Garantidos','Tipo','Producao','Vl.Total Repouso','Total'])
    for d in funcs:
        w.writerow([d['id'], d['nome'], d['engenho'], d['turma'], d['media'],
            d['folgas'], d['faltas'], d['perder'], d['garantidos'], d['tipo'],
            d['producao'], d['valor_folgas'], d['total']])
    return Response(si.getvalue(), mimetype='text/csv',
        headers={'Content-Disposition': 'attachment; filename=dashboard_folgas.csv'})

@app.route('/api/sair')
def api_sair():
    func = request.environ.get('werkzeug.server.shutdown')
    if func:
        func()
    else:
        import os, signal
        os.kill(os.getpid(), signal.SIGINT)
    return 'Encerrando...'

if __name__ == '__main__':
    def abrir_navegador():
        webbrowser.open_new("http://127.0.0.1:5000")
        
    # use_reloader=False evita que o servidor inicie duas vezes e abra duas abas
    threading.Timer(1.2, abrir_navegador).start()
    app.run(debug=True, port=5000, use_reloader=False)
