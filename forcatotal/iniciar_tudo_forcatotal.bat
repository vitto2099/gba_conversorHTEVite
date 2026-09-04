@echo off
chcp 65001 > nul
title FORÇA TOTAL - Executar Pipeline Completo na Máxima Velocidade
color 0C

echo ============================================================================
echo          ⚡ FORÇA TOTAL: PIPELINE COMPLETO NA MÁXIMA VELOCIDADE ⚡
echo ============================================================================
echo.
echo   [1/3] Etapa 1: Conversão Rápida com 8 Threads
echo   [2/3] Etapa 2: Tradução Paralela com 4 Threads
echo   [3/3] Etapa 3: Extração com 8 Threads
echo.
echo ============================================================================
echo.

echo [1/3] Iniciando Conversão Turbo de PDFs...
python "%~dp01_converter_forcatotal.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um erro na Etapa 1.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/3] Iniciando Tradução Paralela de Markdowns...
python "%~dp02_traduzir_forcatotal.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um erro na Etapa 2.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] Iniciando Extração Rápida de Dados...
python "%~dp03_extrair_forcatotal.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um erro na Etapa 3.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ============================================================================
echo          ⚡ PIPELINE FORÇA TOTAL FINALIZADO COM SUCESSO MÁXIMO! ⚡
echo ============================================================================
echo.
echo Pressione qualquer tecla para sair...
pause > nul
