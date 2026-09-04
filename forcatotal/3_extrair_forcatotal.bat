@echo off
chcp 65001 > nul
title FORÇA TOTAL - Etapa 3: Extrator Regulatório Turbo
color 0C

echo ============================================================================
echo      ⚡ FORÇA TOTAL: EXTRATOR REGULATÓRIO E CLÍNICO (TURBO) ⚡
echo ============================================================================
echo  - 8 Threads de mineracao concorrente com Regex compilada em C
echo  - Exportacao instantanea para Excel, JSON e CSV
echo  - Salva na pasta 'dados_extraidos/'
echo ============================================================================
echo.

python "%~dp03_extrair_forcatotal.py" %*

echo.
echo Pressione qualquer tecla para fechar...
pause > nul
