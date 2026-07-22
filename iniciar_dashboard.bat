@echo off
setlocal enabledelayedexpansion
title Sistema Mota - Iniciando...
color 0A

echo.
echo  ====================================================
echo         SISTEMA MOTA - FOLGAS E REPOUSOS
echo  ====================================================
echo.

:: ─────────────────────────────────────────────
:: 1. DETECTAR ARQUITETURA DO WINDOWS
:: ─────────────────────────────────────────────
echo  [1/5] Detectando arquitetura do sistema...
set ARCH=x64
if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    if not defined PROCESSOR_ARCHITEW6432 (
        set ARCH=x86
    )
)
echo       Arquitetura detectada: %ARCH%
echo.

:: ─────────────────────────────────────────────
:: 2. VERIFICAR SE PYTHON JA ESTA INSTALADO
:: ─────────────────────────────────────────────
echo  [2/5] Verificando Python...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
    echo       [OK] !PYVER! encontrado.
    goto :verificar_libs
)

:: Python nao encontrado - fazer download e instalar
echo       [AVISO] Python nao encontrado. Iniciando instalacao automatica...
echo.

:: Definir URL de acordo com a arquitetura
if "%ARCH%"=="x64" (
    set PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
    set PY_FILE=%TEMP%\python_installer_x64.exe
) else (
    set PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9.exe
    set PY_FILE=%TEMP%\python_installer_x86.exe
)

echo       Baixando Python 3.11.9 (%ARCH%) - aguarde...
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_FILE%' -UseBasicParsing }"
if %errorlevel% neq 0 (
    echo.
    echo  [ERRO] Falha ao baixar o Python. Verifique sua conexao com a internet.
    echo         Voce pode instalar manualmente em: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo       Instalando Python (isso pode demorar alguns minutos)...
"%PY_FILE%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
if %errorlevel% neq 0 (
    echo.
    echo  [ERRO] Falha na instalacao do Python. Tente instalar manualmente.
    pause
    exit /b 1
)

:: Atualizar PATH para a sessao atual
for /f "tokens=*" %%p in ('powershell -Command "[System.Environment]::GetEnvironmentVariable(\"PATH\", \"User\")"') do set PATH=%%p;%PATH%

echo       [OK] Python instalado com sucesso!
echo.

:verificar_libs
:: ─────────────────────────────────────────────
:: 3. VERIFICAR E INSTALAR BIBLIOTECAS
:: ─────────────────────────────────────────────
echo  [3/5] Verificando bibliotecas necessarias...

:: Atualizar pip silenciosamente
python -m pip install --upgrade pip --quiet >nul 2>&1

:: Lista de pacotes necessarios
set LIBS=flask xlrd xlwt html2pdf

set LIBS_FALTANDO=0
for %%L in (flask xlrd xlwt) do (
    python -c "import %%L" >nul 2>&1
    if !errorlevel! neq 0 (
        echo       Instalando %%L...
        python -m pip install %%L --quiet
        if !errorlevel! neq 0 (
            echo  [ERRO] Falha ao instalar %%L. Verifique sua conexao.
            set LIBS_FALTANDO=1
        ) else (
            echo       [OK] %%L instalado.
        )
    ) else (
        echo       [OK] %%L ja instalado.
    )
)

if %LIBS_FALTANDO% equ 1 (
    echo.
    echo  [ERRO] Algumas bibliotecas nao puderam ser instaladas.
    pause
    exit /b 1
)
echo.

:: ─────────────────────────────────────────────
:: 4. VERIFICAR PASTA DE DADOS
:: ─────────────────────────────────────────────
echo  [4/5] Verificando pasta de dados...
if not exist "dados\" (
    echo       A pasta dados nao existe. Criando...
    mkdir dados
    echo.
    echo  [AVISO] Pasta dados criada!
    echo          Adicione as planilhas do mes dentro de uma subpasta.
    echo          Exemplo: dados\06-2026\
    echo          Depois execute o programa novamente.
    pause
    exit /b 1
)

:: Verificar se ha pelo menos uma subpasta com planilhas
set PLANILHAS_OK=0
for /d %%D in ("dados\*") do (
    if exist "%%D\*.xls" set PLANILHAS_OK=1
    if exist "%%D\*.xlsx" set PLANILHAS_OK=1
)

:: Verificar tambem planilhas diretamente na raiz de dados
dir /b "dados\*.xls*" >nul 2>&1
if %errorlevel% equ 0 set PLANILHAS_OK=1

if %PLANILHAS_OK% equ 0 (
    echo  [ERRO] Nenhuma planilha encontrada na pasta dados ou subpastas.
    echo         Coloque os arquivos XLS dentro de dados\MM-AAAA\
    echo         Exemplo: dados\06-2026\
    echo.
    pause
    exit /b 1
)
echo       [OK] Planilhas encontradas.
echo.

:: ─────────────────────────────────────────────
:: 5. INICIAR O DASHBOARD
:: ─────────────────────────────────────────────
echo  [5/5] Iniciando servidor...
echo.
echo  ====================================================
echo   Pronto! O navegador abrira automaticamente.
echo   Para encerrar o sistema, feche esta janela.
echo  ====================================================
echo.
cd dashboard
python app.py

pause
