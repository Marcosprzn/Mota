import xlrd
from datetime import datetime, timedelta
from collections import OrderedDict
import os

def encontrar_planilhas():
    diretorio_base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'dados')
    if not os.path.exists(diretorio_base):
        raise FileNotFoundError(f"Diretório {diretorio_base} não encontrado.")
    
    pares = []
    
    pastas_para_olhar = [diretorio_base]
    for p in os.listdir(diretorio_base):
        caminho_p = os.path.join(diretorio_base, p)
        if os.path.isdir(caminho_p):
            pastas_para_olhar.append(caminho_p)
            
    for pasta in set(pastas_para_olhar):
        arq_relatorio = None
        arq_lancamentos = None
        for arquivo in os.listdir(pasta):
            if arquivo.endswith('.xls') or arquivo.endswith('.xlsx'):
                nome_min = arquivo.lower()
                if 'relatorio' in nome_min or 'relatório' in nome_min or 'salario' in nome_min or 'salário' in nome_min:
                    arq_relatorio = os.path.join(pasta, arquivo)
                elif 'lancamento' in nome_min or 'lançamento' in nome_min:
                    arq_lancamentos = os.path.join(pasta, arquivo)
        if arq_relatorio and arq_lancamentos:
            pares.append((arq_relatorio, arq_lancamentos))
            
    if not pares:
        raise FileNotFoundError("Não foi possível encontrar planilhas na pasta 'dados' ou em suas subpastas.")
        
    return pares

def serial_para_data(serial):
    if not serial or serial == '0.0':
        return None
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except ValueError:
        return None

def processar_par(arq_relatorio, arq_lancamentos):
    wb2 = xlrd.open_workbook(arq_lancamentos)
    sheet2 = wb2.sheet_by_index(0)

    turma_por_id = {}
    media_por_id = {}
    folgas_por_id = {}
    faltas_por_id = {}
    descontar_por_id = {}
    meia_por_id = {}
    dias_trab_por_id = {}
    soma_g_meia_por_id = {}

    current_id = None
    eventos_temp = []
    competencia_data = None

    def processa_eventos(eid, eventos, media_val=0.0):
        nonlocal competencia_data

        # ── Encontrar o primeiro dia de trabalho real ──
        primeiro_trabalho = None
        for dt, tipo, horas, colg in eventos:
            if dt is None:
                continue
            if tipo not in ('FOLGA', 'FERIADO', 'FALTA NAO JUSTIFICADA', 'FERIAS', 'AFASTADO') and 'Atestado' not in tipo:
                primeiro_trabalho = dt
                break

        # ── Filtrar: ignorar FOLGA/FERIADO/FALTA antes da admissão ──
        evt_filtrados = []
        for item in eventos:
            dt = item[0]
            if dt is None:
                evt_filtrados.append(item)
                continue
            if primeiro_trabalho and dt < primeiro_trabalho:
                tipo = item[1]
                if tipo in ('FOLGA', 'FERIADO', 'FALTA NAO JUSTIFICADA'):
                    continue
            evt_filtrados.append(item)

        folga_total = 0
        falta_total = 0
        meia = False
        semanas = {}
        g_por_dia = {}

        for dt, tipo, horas, colg in evt_filtrados:
            if dt is None:
                continue
            if competencia_data is None:
                competencia_data = dt.strftime('%m/%Y')
            iso = dt.isocalendar()
            sem = (iso[0], iso[1])
            if sem not in semanas:
                semanas[sem] = {'folga': 0, 'feriado': 0, 'falta': 0}
            if tipo == 'FOLGA':
                semanas[sem]['folga'] += 1
                folga_total += 1
            elif tipo == 'FERIADO':
                semanas[sem]['feriado'] += 1
                folga_total += 1
            elif tipo == 'FALTA NAO JUSTIFICADA':
                semanas[sem]['falta'] += 1
                falta_total += 1

            if dt.weekday() < 5 and tipo not in ('FOLGA', 'FERIADO', 'FALTA NAO JUSTIFICADA', 'FERIAS', 'AFASTADO') and 'Atestado' not in tipo:
                chave = (dt.year, dt.month, dt.day)
                if chave not in g_por_dia:
                    g_por_dia[chave] = 0.0
                g_por_dia[chave] += colg

        if media_val > 0:
            for chave, soma_g in g_por_dia.items():
                if soma_g < 54.80:
                    meia = True

        repousos_perder = 0
        for sem, info in semanas.items():
            repousos_semana = info['folga'] + info['feriado']
            if info['falta'] > 0 and repousos_semana > 0:
                if info['feriado']:
                    repousos_perder += repousos_semana
                else:
                    repousos_perder += min(repousos_semana, info['falta'])

        dias_trab_total = set()
        dias_meia_conj = set()
        for dt, tipo, horas, colg in evt_filtrados:
            if dt is None:
                continue
            chave = (dt.year, dt.month, dt.day)
            if tipo not in ('FOLGA', 'FERIADO', 'FALTA NAO JUSTIFICADA', 'FERIAS', 'AFASTADO') and 'Atestado' not in tipo:
                dias_trab_total.add(chave)
            if dt.weekday() < 5 and tipo not in ('FOLGA', 'FERIADO', 'FALTA NAO JUSTIFICADA', 'FERIAS', 'AFASTADO') and 'Atestado' not in tipo:
                if chave in g_por_dia and g_por_dia[chave] < 54.80:
                    dias_meia_conj.add(chave)
        dias_trab_por_id[eid] = max(1, len(dias_trab_total) - len(dias_meia_conj))
        soma_g_meia_por_id[eid] = sum(g_por_dia[ch] for ch in dias_meia_conj if ch in g_por_dia)

        folgas_por_id[eid] = folga_total
        faltas_por_id[eid] = falta_total
        descontar_por_id[eid] = repousos_perder
        return meia

    for r in range(sheet2.nrows):
        row = [str(sheet2.cell_value(r, c)).strip() for c in range(sheet2.ncols)]
        if 'Funcionário' in row[0] or 'Funcion' in row[0]:
            if current_id and current_id not in folgas_por_id:
                m_val = media_por_id.get(current_id, 0.0)
                meia_val = processa_eventos(current_id, eventos_temp, m_val)
                meia_por_id[current_id] = meia_val
            raw_id = str(sheet2.cell_value(r, 1)).strip()
            try:
                current_id = str(int(float(raw_id))) if '.' in raw_id else raw_id
            except ValueError:
                current_id = raw_id
            turma_raw = row[3].replace('Turma: ', '') if 'Turma:' in row[3] else row[3]
            turma_por_id[current_id] = turma_raw.split(' - ')[0]
            eventos_temp = []
        elif row[0].startswith('M') and 'dia:' in row[0]:
            if current_id:
                try:
                    val_str = row[0].split(':')[1].strip().replace(',', '.')
                    m_val = float(val_str)
                    media_por_id[current_id] = m_val
                except (IndexError, ValueError):
                    m_val = 0.0
                    media_por_id[current_id] = 0.0
                if current_id not in folgas_por_id:
                    meia_val = processa_eventos(current_id, eventos_temp, m_val)
                    meia_por_id[current_id] = meia_val
        elif current_id and len(row) > 2:
            try:
                float(row[0])
                tipo = row[2]
                data = serial_para_data(row[0])
                horas = float(row[4]) if row[4] else 0.0
                colg = float(row[6]) if row[6] else 0.0
                eventos_temp.append((data, tipo, horas, colg))
            except ValueError:
                pass

    if current_id and current_id not in folgas_por_id:
        m_val = media_por_id.get(current_id, 0.0)
        meia_val = processa_eventos(current_id, eventos_temp, m_val)
        meia_por_id[current_id] = meia_val

    wb1 = xlrd.open_workbook(arq_relatorio)
    sheet1 = wb1.sheet_by_index(0)

    resultados = []
    engenho_atual = ''

    for r in range(sheet1.nrows):
        row = [str(sheet1.cell_value(r, c)).strip() for c in range(sheet1.ncols)]
        cabecalho = row[0]
        cab_lower = cabecalho.lower()
        if 'fundo agricola' in cab_lower or 'fundo agrícola' in cab_lower:
            partes = cabecalho.split(' - ', 1)
            engenho_atual = partes[1] if len(partes) > 1 else cabecalho
            continue
        if '.' in cabecalho:
            try:
                cab_clean = str(int(float(cabecalho)))
            except ValueError:
                cab_clean = cabecalho
        else:
            cab_clean = cabecalho
        if cab_clean in ('Total:', 'Total Geral:', '') or not cab_clean.isdigit():
            continue
        func_id = cab_clean
        func_nome = row[1]
        producao = float(row[2]) if row[2] else 0.0
        media = media_por_id.get(func_id, 0.0)
        qtd_folgas = folgas_por_id.get(func_id, 0)
        qtd_faltas = faltas_por_id.get(func_id, 0)
        repousos_perder = descontar_por_id.get(func_id, qtd_faltas)
        turma = turma_por_id.get(func_id, 'Nao encontrado')
        meia = meia_por_id.get(func_id, False)
        folgas_ajustadas = max(0, qtd_folgas - repousos_perder)

        if meia:
            dias_trab = dias_trab_por_id.get(func_id, 0)
            soma_g_meia = soma_g_meia_por_id.get(func_id, 0.0)
            if folgas_ajustadas > 0 and soma_g_meia >= folgas_ajustadas * 27.40:
                base_media = round(producao - 54.80 * folgas_ajustadas * 1.5, 2)
                producao_final = round(producao - 54.80 * folgas_ajustadas, 2)
                media_ajust = round(base_media / dias_trab, 2) if dias_trab > 0 else media
                valor_folgas = round(media_ajust * folgas_ajustadas, 2)
            else:
                folgas_ajustadas = 0
                base_media = round(producao - soma_g_meia, 2)
                media_ajust = round(base_media / dias_trab, 2) if dias_trab > 0 else media
                producao_final = producao
                valor_folgas = 0.0
            soma_total = round(producao_final + valor_folgas, 2)
            tipo = 'MEIA'
        else:
            producao_final = producao
            media_ajust = media
            valor_folgas = round(media * folgas_ajustadas, 2)
            soma_total = round(producao_final + valor_folgas, 2)
            tipo = 'INTEGRAL'

        resultados.append({
            'id': int(func_id),
            'nome': func_nome,
            'engenho': engenho_atual,
            'turma': turma,
            'media': media_ajust,
            'folgas': qtd_folgas,
            'faltas': qtd_faltas,
            'perder': repousos_perder,
            'garantidos': folgas_ajustadas,
            'tipo': tipo,
            'producao': producao_final,
            'valor_folgas': valor_folgas,
            'total': soma_total
        })

    return resultados, competencia_data or 'Desconhecida'

def processar():
    pares = encontrar_planilhas()
    resultados_finais = []
    periodos_disponiveis = set()
    
    for arq_relatorio, arq_lancamentos in pares:
        resultados_par, comp = processar_par(arq_relatorio, arq_lancamentos)
        periodos_disponiveis.add(comp)
        for r in resultados_par:
            r['competencia'] = comp
            resultados_finais.append(r)
            
    lista_periodos = sorted(list(periodos_disponiveis))
    return {
        'funcionarios': resultados_finais,
        'periodos_disponiveis': lista_periodos,
        'competencia': ', '.join(lista_periodos)
    }
