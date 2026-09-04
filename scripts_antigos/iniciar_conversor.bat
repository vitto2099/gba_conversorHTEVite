@echo off
chcp 65001 > nul
title Conversor G-BA PDF para Markdown com Traducao (EN)
color 0B

echo ============================================================================
echo      G-BA PDF para Markdown com Traducao Completa em Ingles
echo ============================================================================
echo  - Protecao de CPU/RAM ativada (PC nao trava)
echo  - Barra de progresso ao vivo no terminal
echo  - Pressione Ctrl+C a qualquer momento para pausar
echo ============================================================================
echo.

python converter_pdf_para_markdown_en.py %*

echo.
pause
