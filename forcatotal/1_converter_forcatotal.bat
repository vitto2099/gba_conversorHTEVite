@echo off
chcp 65001 > nul
title FORÇA TOTAL - Etapa 1: Conversor PDF para Markdown Turbo
color 0C

echo ============================================================================
echo      ⚡ FORÇA TOTAL: CONVERSOR PDF PARA MARKDOWN (TURBO) ⚡
echo ============================================================================
echo  - 8 Workers Paralelos ativados (Core i5 13a Gen 12 threads)
echo  - Latencia zero (sem pausas)
echo  - Prioridade de CPU Alta
echo  - Pressione Ctrl+C a qualquer momento para pausar
echo ============================================================================
echo.

python "%~dp01_converter_forcatotal.py" %*

echo.
echo Pressione qualquer tecla para fechar...
pause > nul
