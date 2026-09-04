@echo off
chcp 65001 > nul
title Etapa 1: Conversor PDF para Markdown (Alemão)
color 0A

echo ============================================================================
echo   ETAPA 1: Conversor de PDF para Markdown (Alemão Original)
echo ============================================================================
echo   - Preserva tabelas e formatação original com PyMuPDF4LLM
echo   - Organiza pastas e nomes de arquivos em Inglês
echo   - Baixo uso de CPU/RAM (o computador não trava)
echo   - Pressione Ctrl+C a qualquer momento para pausar
echo ============================================================================
echo.

python 1_converter_pdf_para_markdown.py %*

echo.
echo Pressione qualquer tecla para fechar esta janela...
pause > nul
