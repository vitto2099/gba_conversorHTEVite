@echo off
chcp 65001 > nul
title Etapa 2: Tradutor de Markdown (Alemão para Inglês)
color 0B

echo ============================================================================
echo   ETAPA 2: Tradutor de Markdown (Alemão para Inglês)
echo ============================================================================
echo   - Traduz parágrafos mantendo estrutura, títulos e tabelas
echo   - Proteção inteligente contra rate-limits e quedas de conexão
echo   - Retoma automaticamente de onde parou
echo   - Pressione Ctrl+C a qualquer momento para pausar
echo ============================================================================
echo.

python 2_traduzir_markdown_en.py %*

echo.
echo Pressione qualquer tecla para fechar esta janela...
pause > nul
