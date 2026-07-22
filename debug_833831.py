import xlrd
from datetime import datetime, timedelta

wb2 = xlrd.open_workbook(r'lançamentos de funcionarios.xls')
sheet2 = wb2.sheet_by_index(0)

def serial_para_data(serial):
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except:
        return None

current_id = None
eventos = []
dias = ['Seg','Ter','Qua','Qui','Sex','Sab','Dom']
for r in range(sheet2.nrows):
    row = [str(sheet2.cell_value(r, c)).strip() for c in range(sheet2.ncols)]
    if 'Funcion' in row[0] and '833831' in row[1]:
        current_id = '833831'
        print('Func:', row[2], 'Turma:', row[3])
        continue
    if current_id and 'Funcion' in row[0]:
        break
    if current_id and len(row) > 2:
        try:
            float(row[0])
            dt = serial_para_data(row[0])
            tipo = row[2]
            horas = float(row[4]) if row[4] else 0.0
            colg = float(row[6]) if row[6] else 0.0
            eventos.append((dt, tipo, horas, colg))
            if dt:
                ds = dias[dt.weekday()]
                print('  ' + ds + ' ' + str(dt.day).zfill(2) + '/' + str(dt.month).zfill(2) + '  ' + tipo[:35] + ' E=' + str(horas) + ' G=' + str(colg))
        except:
            pass

print()
semanas = {}
for dt, tipo, horas, colg in eventos:
    if dt is None: continue
    iso = dt.isocalendar()
    sem = (iso[0], iso[1])
    if sem not in semanas:
        semanas[sem] = {'folga':0,'feriado':0,'falta':0,'dias':[]}
    if tipo == 'FOLGA':
        semanas[sem]['folga'] += 1
        semanas[sem]['dias'].append('FOLGA ' + str(dt.day)+'/'+str(dt.month))
    elif tipo == 'FERIADO':
        semanas[sem]['feriado'] += 1
        semanas[sem]['dias'].append('FERIADO ' + str(dt.day)+'/'+str(dt.month))
    elif tipo == 'FALTA NAO JUSTIFICADA':
        semanas[sem]['falta'] += 1
        semanas[sem]['dias'].append('FALTA ' + str(dt.day)+'/'+str(dt.month))

print('Semanas:')
for sem, info in sorted(semanas.items()):
    perder = 0
    rep = info['folga'] + info['feriado']
    if info['falta'] > 0 and rep > 0:
        if info['feriado']:
            perder = rep
        else:
            perder = min(rep, info['falta'])
    linha = '  Sem ' + str(sem[0]) + '-' + str(sem[1]).zfill(2) + ': F=' + str(info['folga']) + ' Fe=' + str(info['feriado']) + ' Fa=' + str(info['falta']) + '  Perder=' + str(perder) + '  ' + str(info['dias'])
    print(linha)

g_por_dia = {}
for dt, tipo, horas, colg in eventos:
    if dt is None: continue
    if dt.weekday() < 5 and tipo not in ('FOLGA','FERIADO','FALTA NAO JUSTIFICADA','FERIAS','AFASTADO') and 'Atestado' not in tipo:
        chave = (dt.year, dt.month, dt.day)
        if chave not in g_por_dia:
            g_por_dia[chave] = 0.0
        g_por_dia[chave] += colg

print()
print('Dias com G:')
dias_meia = 0
total_dias = 0
for chave, soma in sorted(g_por_dia.items()):
    total_dias += 1
    m = ' *MEIA*' if soma < 54.80 else ''
    if soma < 54.80:
        dias_meia += 1
    print('  ' + str(chave[2]).zfill(2) + '/' + str(chave[1]).zfill(2) + ' G=' + str(soma) + m)
print('Total dias: ' + str(total_dias) + ', MEIA: ' + str(dias_meia) + ', efetivos: ' + str(total_dias - dias_meia))
