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
        wb = Workbook()
        ws = wb.add_sheet('Dados')
        cab = ['ID','Nome','Engenho','Turma','Vl.Repouso','Folgas','Faltas','Perder','Garantidos','Tipo','Producao','Vl.Total Repouso','Total']
        estilo_cab = easyxf('font: bold on, colour white; pattern: pattern solid, fore_colour dark_blue; borders: left thin, right thin, top thin, bottom thin;')
        estilo_num = easyxf('font: name Calibri, height 180; num_format: #,##0.00; borders: left thin, right thin, top thin, bottom thin;')
        estilo_int = easyxf('font: name Calibri, height 180; num_format: #,##0; borders: left thin, right thin, top thin, bottom thin;')
        estilo_txt = easyxf('font: name Calibri, height 180; borders: left thin, right thin, top thin, bottom thin;')
        for c, t in enumerate(cab):
            ws.write(0, c, t, estilo_cab)
            ws.col(c).width = int(([10,40,22,8,10,8,8,8,10,8,16,16,16][c]) * 256)
        for i, d in enumerate(funcs):
            r = i + 1
            ws.write(r, 0, int(d['id']), estilo_int)
            ws.write(r, 1, d['nome'], estilo_txt)
            ws.write(r, 2, d['engenho'], estilo_txt)
            ws.write(r, 3, d['turma'], estilo_txt)
            ws.write(r, 4, d['media'], estilo_num)
            ws.write(r, 5, d['folgas'], estilo_int)
            ws.write(r, 6, d['faltas'], estilo_int)
            ws.write(r, 7, d['perder'], estilo_int)
            ws.write(r, 8, d['garantidos'], estilo_int)
            ws.write(r, 9, d['tipo'], estilo_txt)
            ws.write(r, 10, d['producao'], estilo_num)
            ws.write(r, 11, d['valor_folgas'], estilo_num)
            ws.write(r, 12, d['total'], estilo_num)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return Response(buf.getvalue(), mimetype='application/vnd.ms-excel',
            headers={'Content-Disposition': 'attachment; filename=dashboard_folgas.xls'})
    except ImportError:
        pass

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
