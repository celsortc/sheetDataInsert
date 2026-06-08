@echo off
title Gerando .exe - Processador de Estagiarios

REM Detecta automaticamente o Python instalado
for /f "delims=" %%i in ('where /r "%LOCALAPPDATA%\Python" python.exe 2^>nul') do set PYTHON=%%i
if not defined PYTHON (
    for /f "delims=" %%i in ('where /r "%APPDATA%\Python" python.exe 2^>nul') do set PYTHON=%%i
)
if not defined PYTHON (
    for /f "delims=" %%i in ('where python.exe 2^>nul') do set PYTHON=%%i
)

if not defined PYTHON (
    echo ERRO: Python nao encontrado no computador.
    pause
    exit /b 1
)

echo  Python encontrado: %PYTHON%
echo.
echo  Instalando dependencias...
"%PYTHON%" -m pip install pdfplumber openpyxl pyinstaller --quiet

echo.
echo  Compilando o executavel...
"%PYTHON%" -m PyInstaller --onefile --windowed --name "Processador_Estagiarios" leitor_estagiarios.py

echo.
if exist "dist\Processador_Estagiarios.exe" (
    echo  ============================================
    echo   SUCESSO! .exe gerado em: dist\
    echo  ============================================
    explorer dist
) else (
    echo  Algo deu errado. Veja as mensagens acima.
)
pause
