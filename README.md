# Pipeline Modular de Processamento e Análise de Documentos G-BA 🇩🇪 ➔ 🇬🇧

Este projeto implementa um pipeline desacoplado e eficiente em **3 etapas independentes** para converter, traduzir e extrair informações regulatórias e clínicas do acervo de decisões e dossiês de avaliação de tecnologias em saúde (HTA) da Alemanha (**G-BA / Gemeinsamer Bundesausschuss** e **IQWiG**).

---

## 🚀 Por que o pipeline foi dividido em 3 etapas?

Originalmente, tentar extrair PDFs e traduzir centenas de páginas via requisições de rede ao mesmo tempo tornava o processo lento para um acervo de mais de 30 mil PDFs (arquivos com 350 páginas levavam vários minutos apenas para a rede responder).

A divisão em 3 scripts resolve isso:
1. **Conversão instantânea (Alemão):** Converte milhares de PDFs para Markdown no idioma original em segundos por arquivo, preservando tabelas e títulos perfeitamente com `PyMuPDF4LLM`.
2. **Tradução desacoplada (Inglês):** Traduz os arquivos `.md` parágrafo por parágrafo com controle de taxa de requisições, retentativas automáticas e retomada de onde parou.
3. **Extração de inteligência regulatória:** Minera os campos mais importantes de cada documento (Princípio Ativo, Indicação, Terapia Comparadora, Grau de Benefício Adicional *Zusatznutzen*, etc.) e compila tudo em **Excel**, **JSON** e **CSV**.

---

## 🛠️ Requisitos e Instalação

Certifique-se de estar com o Python 3.10+ instalado e execute no terminal:

```bash
pip install pymupdf pymupdf4llm deep-translator rich openpyxl pandas psutil
```

---

## 📋 Como Usar Cada Etapa

### Etapa 1: Converter PDF para Markdown (Alemão Original)
Converte os PDFs da pasta `gba/` para Markdown em `gba_markdown_de/`, renomeando as pastas e arquivos para Inglês com base na planilha `nomes_pastas_arquivos_traduzidos.xlsx`.

```bash
# Execução padrão (rápida, 2 threads, prioridade baixa no Windows):
python 1_converter_pdf_para_markdown.py

# Modo Ultra-Leve (1 thread, prioridade ociosa, ventiladores 100% silenciosos):
python 1_converter_pdf_para_markdown.py --eco

# Mais rápido (para processadores potentes com 4 threads):
python 1_converter_pdf_para_markdown.py --workers 4

# Testar com apenas 5 arquivos:
python 1_converter_pdf_para_markdown.py --limit 5

# Filtrar por procedimento específico (ex: 1219 ou Pembrolizumab):
python 1_converter_pdf_para_markdown.py --filter 1219
```

> **Características da Etapa 1:**
> - Processa PDFs em fatias de páginas (lotes de 25 páginas), mantendo o uso de RAM sempre baixo (< 250 MB).
> - Gravação atômica (`.tmp` ➔ `.md`) para evitar corrupção caso o processo seja interrompido.
> - Pula automaticamente arquivos que já foram convertidos.

---

### Etapa 2: Traduzir Markdown para Inglês
Lê os arquivos da pasta `gba_markdown_de/` e grava as versões traduzidas em `gba_markdown_en/`.

```bash
# Execução padrão:
python 2_traduzir_markdown_en.py

# Modo suave (com pausas maiores anti-bloqueio):
python 2_traduzir_markdown_en.py --eco

# Testar com apenas 2 arquivos:
python 2_traduzir_markdown_en.py --limit 2

# Filtrar por arquivo ou procedimento específico:
python 2_traduzir_markdown_en.py --filter 1219
```

> **Características da Etapa 2:**
> - Segmentação em blocos de parágrafos (`\n\n` <= 3200 caracteres), mantendo títulos Markdown e tabelas íntegras.
> - Tolerância a falhas de rede (retentativas com backoff exponencial se houver erro 429 ou queda de conexão).
> - Pula arquivos que já foram traduzidos anteriormente.

---

### Etapa 3: Extrair Dados Estruturados
Varre os arquivos Markdown e extrai os principais dados clínicos e regulatórios, salvando os resultados em `dados_extraidos_gba.xlsx`, `dados_extraidos_gba.json` e `dados_extraidos_gba.csv`.

```bash
# Executar na pasta de markdowns (detecta automaticamente gba_markdown_de ou gba_markdown_en):
python 3_extrair_dados_markdown.py

# Especificar pasta de origem explicitamente:
python 3_extrair_dados_markdown.py --source gba_markdown_de

# Definir nome personalizado para os relatórios:
python 3_extrair_dados_markdown.py --output-prefix meu_relatorio_gba

# Testar em lote menor:
python 3_extrair_dados_markdown.py --limit 50
```

#### Campos Extraídos:
- **`id_procedimento`**: Identificador do processo G-BA (ex: `D-1219`).
- **`principio_ativo`**: Substância ativa (*Wirkstoff* / Active Substance).
- **`nome_comercial`**: Nome da marca (*Handelsname* / Trade Name).
- **`tipo_documento`**: Decisão (Beschluss), Razões Determinantes (Tragende Gründe), Dossiê Módulo 1 a 5, Relatório IQWiG, etc.
- **`beneficio_adicional`**: Classificação oficial do benefício (*Beträchtlich*, *Gering*, *Nicht quantifizierbar*, *Kein Zusatznutzen*).
- **`status_orphan_drug`**: Se é medicamento para doenças raras.
- **`indicacao_terapeutica`**: Resumo da indicação clínica aprovada.
- **`terapia_comparadora`**: Terapia comparadora apropriada (*Zweckmäßige Vergleichstherapie* - ZVT).
- **`data_decisao`**: Data de publicação/deliberação do G-BA.

---

### Execução Rápida via Arquivos .BAT (Dois Cliques)

Para facilitar no Windows, cada etapa possui seu próprio arquivo `.bat` pronto para execução com duplo clique no Windows Explorer:

- 🚀 **`menu.bat`**: **Menu interativo** para escolher qualquer etapa (1, 2, 3), executar o pipeline completo ou modo ECO sem digitar comandos.
- 📄 **`1_converter_pdf_para_markdown.bat`**: Inicia a conversão dos PDFs para Markdown em Alemão.
- 🌐 **`2_traduzir_markdown_en.bat`**: Inicia a tradução dos Markdowns para Inglês.
- 📊 **`3_extrair_dados_markdown.bat`**: Inicia a extração e salva as planilhas e relatórios em `dados_extraidos/`.

---

## ⚡ Modo Especial: FORÇA TOTAL (`forcatotal/`)

Para quem deseja **o menor tempo possível de processamento** e tem um processador multi-core com boa memória RAM (como o **Intel Core i5 de 13ª Geração com 12 threads e 24 GB de RAM**), criamos a pasta dedicada **`forcatotal/`**.

> [!TIP]
> **Quando usar o Modo Força Total?**
> - **Recomendado para SSD Interno (ex: `C:`):** Quando você copiar os arquivos para o SSD, use o Força Total para aproveitar 100% da velocidade sem gargalo de USB/HD mecânico.
> - **Se estiver no HD Externo:** O modo convencional da raiz ainda é mais silencioso e seguro termicamente, mas o Força Total também pode ser usado se você quiser prioridade alta de CPU.

### 🏎️ Comparativo: Padrão vs. Força Total

| Recurso | Versão Convencional (Raiz) | ⚡ Versão FORÇA TOTAL (`forcatotal/`) |
| :--- | :--- | :--- |
| **Threads de Conversão (Etapa 1)** | 2 threads (1 no Eco) | **8 threads paralelas simultâneas** |
| **Tradução (Etapa 2)** | 1 arquivo por vez | **4 arquivos traduzidos ao mesmo tempo** |
| **Extração (Etapa 3)** | 1 thread sequencial | **8 threads paralelas com Regex compilado em C** |
| **Prioridade no Windows** | Baixa / Ociosa (não esquentar) | **Acima do Normal / Máxima CPU** |
| **Pausas / Delays** | Pausas térmicas ativadas | **Latência zero (máximo throughput)** |
| **Execução Completa (1 clique)** | Manual via menu | **`iniciar_tudo_forcatotal.bat` (faz 1 ➔ 2 ➔ 3 direto)** |

### 🚀 Como Usar o Força Total:
1. Abra a pasta **`forcatotal/`**.
2. Dê dois cliques em **`iniciar_tudo_forcatotal.bat`** para rodar todo o pipeline na velocidade máxima.
3. Ou use os executáveis individuais da pasta:
   - `1_converter_forcatotal.bat`: Apenas conversão de PDF para Markdown.
   - `2_traduzir_forcatotal.bat`: Apenas tradução para Inglês.
   - `3_extrair_forcatotal.bat`: Apenas mineração de dados para Excel/JSON/CSV.

---

## 🛡️ Proteção do Computador (Modo Padrão)

Na versão padrão da raiz, os scripts foram desenvolvidos com salvaguardas nativas:
- **Prioridade Reduzida no Windows (`psutil`):** Os scripts rodam com `BELOW_NORMAL_PRIORITY_CLASS` (ou `IDLE_PRIORITY_CLASS` no modo `--eco`). Isso significa que o Windows sempre prioriza o seu navegador, vídeos, reuniões ou jogos; o script usa apenas o processamento ocioso.
- **Consumo de Memória Travado:** Limpeza de lixo (`gc.collect()`) a cada arquivo e leitura sob demanda evitam estouro de RAM mesmo com centenas de páginas.
- **Interface Rich no Terminal:** Exibe porcentagem, arquivos concluídos/restantes, tempo decorrido, estimativa de conclusão (ETA) e monitoramento de CPU e memória RAM em tempo real.

---

## 📁 Estrutura Organizada do Projeto

```
d:\Documentos Alemanha/
│
├── menu.bat                            # Menu interativo principal (inicia tudo)
├── 1_converter_pdf_para_markdown.bat   # Launcher duplo clique da Etapa 1
├── 1_converter_pdf_para_markdown.py    # Script Etapa 1: PDF -> Markdown (Alemão)
├── 2_traduzir_markdown_en.bat          # Launcher duplo clique da Etapa 2
├── 2_traduzir_markdown_en.py           # Script Etapa 2: Markdown Alemão -> Inglês
├── 3_extrair_dados_markdown.bat        # Launcher duplo clique da Etapa 3
├── 3_extrair_dados_markdown.py         # Script Etapa 3: Extrator Regulatório
│
├── forcatotal/                         # ⚡ SUÍTE TURBO (Máxima Velocidade / SSD):
│   ├── iniciar_tudo_forcatotal.bat     # Executa 1 -> 2 -> 3 direto na máxima velocidade
│   ├── 1_converter_forcatotal.bat      # Etapa 1 com 8 threads
│   ├── 1_converter_forcatotal.py
│   ├── 2_traduzir_forcatotal.bat      # Etapa 2 com 4 threads concorrentes
│   ├── 2_traduzir_forcatotal.py
│   ├── 3_extrair_forcatotal.bat        # Etapa 3 com 8 threads e Regex em C
│   ├── 3_extrair_forcatotal.py
│   └── README_FORCATOTAL.md            # Manual do modo turbo
│
├── README.md                           # Este manual de instruções principal
├── .gitignore                          # Ignora PDFs e Markdowns (mantém o README)
├── nomes_pastas_arquivos_traduzidos.xlsx # Mapeamento de pastas/arquivos para Inglês
├── nomes_pastas_arquivos_traduzidos.json # Cache de inicialização instantânea
│
├── gba/                                # PDFs originais baixados (~33 mil arquivos)
├── gba_markdown_de/                    # Markdowns originais em Alemão (Gerado Etapa 1)
├── gba_markdown_en/                    # Markdowns traduzidos em Inglês (Gerado Etapa 2)
│
├── dados_extraidos/                    # Pasta com as saídas geradas na Etapa 3:
│   ├── dados_extraidos_gba.xlsx        # Planilha Excel com colunas estruturadas
│   ├── dados_extraidos_gba.json        # JSON estruturado para integrações
│   └── dados_extraidos_gba.csv         # CSV para análise rápida de dados
│
└── scripts_antigos/                    # Arquivos de download e versões antigas arquivadas
```
