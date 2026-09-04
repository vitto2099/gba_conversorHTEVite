@echo off
chcp 65001 > nul
title FORÇA TOTAL - Etapa 2: Tradutor de Markdown Multi-Thread
color 0C

echo ============================================================================
echo      ⚡ FORÇA TOTAL: TRADUTOR DE MARKDOWN MULTI-THREAD ⚡
echo ============================================================================
echo  - Tradução Concorrente com 4 Threads simultaneas
echo  - Latencia minima entre requisicoes
echo  - Backoff inteligente individual
echo  - Pressione Ctrl+C a qualquer momento para pausar
echo ============================================================================
echo.

python "%~dp02_traduzir_forcatotal.py" %*

echo.
echo Pressione qualquer tecla para fechar...
pause > nul
