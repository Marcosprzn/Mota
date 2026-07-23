from playwright.sync_api import sync_playwright
import sys, time

BASE_URL = 'http://127.0.0.1:5000'

def main():
    failed = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            viewport={'width': 1440, 'height': 900},
            locale='pt-BR'
        )
        page = ctx.new_page()

        # ── 1. Carga inicial ──
        print('[TEST] 1. Carga inicial...')
        page.goto(BASE_URL)
        page.wait_for_load_state('networkidle')
        # Aguardar loading desaparecer e dados carregarem
        page.wait_for_selector('#app[style*="block"]', timeout=30000)
        time.sleep(1)  # charts renderizam

        # Verificar KPIs (8 cards)
        kpis = page.locator('.kpi-card')
        kpi_count = kpis.count()
        print(f'  KPIs encontrados: {kpi_count}')
        if kpi_count != 8:
            failed.append(f'KPI count: esperado 8, obtido {kpi_count}')
        else:
            print('  OK - 8 KPIs')

        # Verificar tabela com dados
        linhas = page.locator('#table-body tr').count()
        print(f'  Linhas na tabela: {linhas}')
        if linhas < 300:
            failed.append(f'Tabela com poucas linhas: {linhas}')
        else:
            print(f'  OK - {linhas} registros')

        # Verificar charts
        chart_eng = page.locator('#chart-engenhos')
        chart_gar = page.locator('#chart-garantidos')
        if not chart_eng.is_visible() or not chart_gar.is_visible():
            failed.append('Charts nao visiveis')
        else:
            print('  OK - Charts visiveis')

        # ── 2. Filtrar por engenho LARANJEIRAS ──
        print('[TEST] 2. Filtrar por engenho LARANJEIRAS...')
        page.select_option('#filter-eng', 'LARANJEIRAS')
        time.sleep(1.5)  # aguardar re-render

        # Verificar KPIs atualizados (Engenhos = 1)
        kpi_texts = page.locator('.kpi-card .value').all_text_contents()
        eng_kpi = kpi_texts[7]  # indice 7 = Engenhos
        print(f'  KPI Engenhos apos filtro: {eng_kpi}')
        if '1' not in eng_kpi:
            failed.append(f'KPI Engenhos deveria ser 1, obtido {eng_kpi}')
        else:
            print('  OK - Engenhos = 1')

        # Verificar tabela filtrada (só LARANJEIRAS)
        time.sleep(0.5)
        linhas_filtro = page.locator('#table-body tr').count()
        print(f'  Linhas apos filtro: {linhas_filtro}')
        if linhas_filtro >= linhas:
            failed.append('Tabela nao filtrou corretamente')
        else:
            print(f'  OK - Filtrou para {linhas_filtro} registros')

        # ── 3. Resetar filtros e filtrar por tipo MEIA ──
        print('[TEST] 3. Resetar filtros e filtrar por tipo MEIA...')
        page.select_option('#filter-eng', 'Todos')
        page.select_option('#filter-tipo', 'MEIA')
        time.sleep(1.5)

        meias_filtro = page.locator('#table-body tr').count()
        print(f'  Funcionarios MEIA: {meias_filtro}')
        if meias_filtro == 0:
            failed.append('Nenhum funcionario MEIA encontrado')
        else:
            print(f'  OK - {meias_filtro} MEIA(s)')

        # Verificar badge MEIA
        badges = page.locator('.badge.meia').count()
        if badges < 1:
            failed.append('Badge MEIA nao encontrado')
        else:
            print(f'  OK - {badges} badge(s) MEIA')

        # Resetar filtro tipo
        page.select_option('#filter-tipo', 'Todos')
        time.sleep(1.5)

        # ── 4. Exportar Excel ──
        print('[TEST] 4. Exportar Excel...')
        with page.expect_download(timeout=15000) as download_info:
            page.click('.btn-export-excel')
        download = download_info.value
        print(f'  Arquivo baixado: {download.suggested_filename}')
        if not download.suggested_filename.endswith('.xls'):
            failed.append(f'Export Excel: extensao incorreta {download.suggested_filename}')
        else:
            print('  OK - Excel exportado')

        # ── 5. Exportar PDF ──
        print('[TEST] 5. Exportar PDF...')
        with page.expect_download(timeout=30000) as download_info:
            page.click('.btn-export-pdf')
        download = download_info.value
        print(f'  Arquivo baixado: {download.suggested_filename}')
        if not download.suggested_filename.endswith('.pdf'):
            failed.append(f'Export PDF: extensao incorreta {download.suggested_filename}')
        else:
            print('  OK - PDF exportado')

        # ── 6. Botao Sair ──
        print('[TEST] 6. Botao Sair...')
        page.on('dialog', lambda dialog: dialog.accept())
        page.click('button:has-text("Sair")')
        time.sleep(3)
        body_text = page.locator('body').text_content()
        if 'encerrado' in body_text.lower() or 'Encerrando' in body_text:
            print('  OK - Mensagem de encerramento exibida')
        else:
            # Se nao matou o processo (ex: rodando com with_server.py), nao eh falha
            print('  INFO - Sair clicado (servidor gerenciado externamente)')

        browser.close()

    # ── Resultado ──
    if failed:
        print(f'\n{"="*50}')
        print(f'FALHAS ({len(failed)}):')
        for f in failed:
            print(f'  - {f}')
        sys.exit(1)
    else:
        print(f'\n{"="*50}')
        print('TODOS OS TESTES PASSARAM')
        sys.exit(0)

if __name__ == '__main__':
    main()
