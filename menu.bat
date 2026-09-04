@echo off
chcp 65001 > nul
title Pipeline G-BA - Menu Principal
color 0F

:MENU
cls
echo ============================================================================
echo                     PIPELINE G-BA - MENU PRINCIPAL
echo ============================================================================
echo.
echo   [1] Etapa 1: Converter PDFs para Markdown (Alemão original)
echo   [2] Etapa 2: Traduzir Markdowns para Inglês
echo   [3] Etapa 3: Extrair Dados Estruturados (Excel / JSON / CSV)
echo.
echo   [4] Executar Pipeline Completo (1 -> 2 -> 3)
echo   [5] Executar Etapa 1 em Modo ECO (1 thread, prioridade ociosa)
echo.
echo   [0] Sair
echo.
echo ============================================================================
set /p OPCAO=" Escolha uma opção [0-5]: "

if "%OPCAO%"=="1" goto ETAPA1
if "%OPCAO%"=="2" goto ETAPA2
if "%OPCAO%"=="3" goto ETAPA3
if "%OPCAO%"=="4" goto COMPLETO
if "%OPCAO%"=="5" goto ETAPA1_ECO
if "%OPCAO%"=="0" goto SAIR

echo.
echo Opção inválida! Tente novamente.
timeout /t 2 > nul
goto MENU

:ETAPA1
cls
call 1_converter_pdf_para_markdown.bat
goto MENU

:ETAPA2
cls
call 2_traduzir_markdown_en.bat
goto MENU

:ETAPA3
cls
call 3_extrair_dados_markdown.bat
goto MENU

:COMPLETO
cls
echo ============================================================================
echo              EXECUTANDO PIPELINE COMPLETO (1 -> 2 -> 3)
echo ============================================================================
echo.
echo [1/3] Iniciando Conversão de PDFs...
python 1_converter_pdf_para_markdown.py
echo.
echo [2/3] Iniciando Tradução para Inglês...
python 2_traduzir_markdown_en.py
echo.
echo [3/3] Iniciando Extração de Dados...
python 3_extrair_dados_markdown.py
echo.
echo ============================================================================
echo              PIPELINE COMPLETO FINALIZADO COM SUCESSO!
echo ============================================================================
pause
goto MENU

:ETAPA1_ECO
cls
call 1_converter_pdf_para_markdown.bat --eco
goto MENU

:SAIR
exit /b 0
