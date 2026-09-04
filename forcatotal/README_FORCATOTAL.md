# ⚡ Modo FORÇA TOTAL - Máxima Velocidade ⚡

Esta pasta contém a versão **Turbo** de altíssimo desempenho do pipeline G-BA, calibrada especificamente para aproveitar a capacidade total do processador (**Intel Core i5-13420H com 12 threads** e **24 GB de RAM**).

---

## 🏎️ Diferenças do Modo Força Total

| Característica | Versão Padrão (Raiz) | ⚡ Versão FORÇA TOTAL (`forcatotal/`) |
| :--- | :--- | :--- |
| **Threads na Conversão (Etapa 1)** | 2 threads (ou 1 no Eco) | **8 threads paralelas** |
| **Tradução (Etapa 2)** | 1 thread sequencial com delay | **4 threads concorrentes** de arquivos |
| **Extração (Etapa 3)** | 1 thread | **8 threads paralelas com Regex em C** |
| **Prioridade de Processo** | Baixa / Ociosa (não esquentar) | **Acima do Normal / Máxima CPU** |
| **Pausas / Delays** | Ativadas para silêncio térmico | **Latência zero (máximo throughput)** |
| **Ideal para** | Segundo plano no HD externo | **SSD interno com foco em terminar rápido** |

---

## 🚀 Como Usar

Você pode rodar os arquivos com dois cliques ou via linha de comando:

### 1. Executar Tudo de uma vez
- Dê dois cliques em **`iniciar_tudo_forcatotal.bat`** (executa 1 ➔ 2 ➔ 3 sequencialmente no modo turbo).

### 2. Executar Etapas Individuais
- **`1_converter_forcatotal.bat`**: Converte todos os PDFs com 8 threads simultâneas.
- **`2_traduzir_forcatotal.bat`**: Traduz os Markdowns com 4 threads simultâneas.
- **`3_extrair_forcatotal.bat`**: Minera os dados clínicos e regulatórios com 8 threads.
