@echo off
chcp 65001 > nul
title Etapa 3: Extrator de Dados Regulatórios e Clínicos
color 0E

echo ============================================================================
echo   ETAPA 3: Extrator de Dados Regulatórios e Clínicos (G-BA)
echo ============================================================================
echo   - Minera Princípio Ativo, Indicação, ZVT, Zusatznutzen, datas, etc.
echo   - Exporta para Excel (.xlsx), JSON (.json) e CSV (.csv)
echo   - Salva tudo na pasta 'dados_extraidos/'
echo ============================================================================
echo.

python 3_extrair_dados_markdown.py %*

echo.
echo Pressione qualquer tecla para fechar esta janela...
pause > nul
