@echo off
setlocal enabledelayedexpansion
title Atualizar Dashboard Farmacia Lider

echo ============================================================
echo   Dashboard Farmacia Lider - Atualizacao Automatica
echo ============================================================
echo.

:: ── Diretório do projeto ─────────────────────────────────────────────────────
set "PROJETO=C:\Users\TI\OneDrive\Trabalho\Vital Contabilidade\Farmacia Lider"
cd /d "%PROJETO%"
if errorlevel 1 (
    echo [ERRO] Pasta do projeto nao encontrada:
    echo        %PROJETO%
    echo.
    echo Verifique se o OneDrive esta sincronizado e o caminho esta correto.
    pause
    exit /b 1
)

:: ── Detectar Python disponível ───────────────────────────────────────────────
echo [1/4] Detectando Python...
set "PYTHON="

if exist "C:\Users\TI\anaconda3\python.exe" (
    set "PYTHON=C:\Users\TI\anaconda3\python.exe"
    goto :python_found
)
if exist "%USERPROFILE%\AppData\Local\anaconda3\python.exe" (
    set "PYTHON=%USERPROFILE%\AppData\Local\anaconda3\python.exe"
    goto :python_found
)
where python >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=python"
    goto :python_found
)
where py >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=py"
    goto :python_found
)

echo [ERRO] Python nao encontrado!
echo Instale o Python em https://www.python.org ou o Anaconda em https://anaconda.com
pause
exit /b 1

:python_found
echo [OK] Python encontrado: %PYTHON%
echo.

:: ── Verificar dependências ────────────────────────────────────────────────────
echo [2/4] Verificando dependencias Python...
%PYTHON% -c "import pandas, openpyxl, git" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Instalando dependencias necessarias...
    %PYTHON% -m pip install pandas openpyxl gitpython --quiet
    if errorlevel 1 (
        echo [AVISO] Falha ao instalar dependencias. Tentando continuar mesmo assim...
    ) else (
        echo [OK] Dependencias instaladas.
    )
) else (
    echo [OK] Dependencias OK.
)
echo.

:: ── Executar script Python ────────────────────────────────────────────────────
echo [3/4] Atualizando dados do dashboard...
echo.
%PYTHON% atualizar_lider.py
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao executar atualizar_lider.py
    echo        Verifique se a planilha Resumo_Farmacia_Lider.xlsx esta fechada no Excel.
    pause
    exit /b 1
)

:: ── Verificar envio ──────────────────────────────────────────────────────────
echo.
echo [4/4] Verificando envio para o servidor...
git status --short >nul 2>&1
if errorlevel 1 (
    echo [AVISO] Git nao encontrado ou nao inicializado.
) else (
    echo [OK] Git operacional.
)

echo.
echo ============================================================
echo   Dashboard atualizado com sucesso!
echo   Acesse o link do Vercel para ver as alteracoes.
echo ============================================================
echo.
pause
