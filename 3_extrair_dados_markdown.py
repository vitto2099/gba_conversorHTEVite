#!/usr/bin/env python3
"""
Etapa 3: Extrator de Dados Regulatórios e Clínicos dos Markdowns
==============================================================
- Minera informações cruciais dos documentos G-BA (Alemão ou Inglês):
    * ID do Procedimento (ex: D-1219)
    * Princípio Ativo (Wirkstoff / Active Substance)
    * Nome Comercial (Handelsname / Brand Name)
    * Indicação Terapêutica (Anwendungsgebiet / Indication)
    * Terapia Comparadora Apropriada (Zweckmäßige Vergleichstherapie / ZVT)
    * Benefício Adicional Concedido (Zusatznutzen: Beträchtlich, Gering, Não quantificável, Sem benefício)
    * Status de Medicamento Órfão (Orphan Drug)
    * Tipo de Documento (Dossiê Módulo 1-5, Beschluss, Tragende Gründe, IQWiG Bericht, etc.)
    * Data da Decisão / Publicação
- Exporta os dados consolidados em:
    * dados_extraidos_gba.xlsx (Planilha formatada)
    * dados_extraidos_gba.json (JSON estruturado completo)
    * dados_extraidos_gba.csv (CSV pronto para análise)
- Exibe resumo estatístico e tabela rica no terminal com Rich.
"""

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.table import Table

# Forçar UTF-8 no terminal Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass

console = Console()


def clean_text_snippet(text: str, max_chars: int = 400) -> str:
    """Limpa quebras de linha e espaços múltiplos em snippets de texto."""
    if not text:
        return ""
    cleaned = re.sub(r"[\r\n\t]+", " ", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    # Remove marcações markdown excedentes
    cleaned = re.sub(r"[*#_`|]", "", cleaned).strip()
    if len(cleaned) > max_chars:
        return cleaned[: max_chars - 3] + "..."
    return cleaned


def extract_procedure_id(path: Path) -> str:
    """Extrai o identificador do procedimento (ex: D-1219, 1219, etc.) pelo caminho."""
    parts = path.parts
    for p in parts:
        m = re.search(r"\b([A-Za-z]?-\d{3,5})\b", p)
        if m:
            return m.group(1).upper()
        m_num = re.search(r"\b(\d{3,5})\b", p)
        if m_num:
            return f"D-{m_num.group(1)}"
    return "N/A"


def identify_document_type(file_path: Path, content_sample: str) -> str:
    """Identifica o tipo de documento (Dossiê, Decisão, Relatório IQWiG, etc.)."""
    path_str = str(file_path).casefold()
    sample = content_sample.casefold()

    if "beschluss" in path_str or "decision" in path_str:
        return "Decisão (Beschluss)"
    if "tragende gruende" in path_str or "tragende gründe" in path_str or "supporting reasons" in path_str:
        return "Razões Determinantes (Tragende Gründe)"
    if "modul 1" in path_str or "module 1" in path_str:
        return "Dossiê - Módulo 1 (Geral)"
    if "modul 2" in path_str or "module 2" in path_str:
        return "Dossiê - Módulo 2 (Resumo Clínico)"
    if "modul 3" in path_str or "module 3" in path_str:
        return "Dossiê - Módulo 3 (Eficácia/Segurança)"
    if "modul 4" in path_str or "module 4" in path_str:
        return "Dossiê - Módulo 4 (Custos/Impacto)"
    if "modul 5" in path_str or "module 5" in path_str:
        return "Dossiê - Módulo 5 (Anexos)"
    if "iqwig" in path_str or "abschlussbericht" in path_str or "final report" in path_str:
        return "Avaliação IQWiG (Bericht)"
    if "stellungnahme" in path_str or "statement" in path_str:
        return "Manifestação (Stellungnahme)"

    # Tenta pelo conteúdo
    if "beschluss des gemeinsamen bundesausschusses" in sample or "decision of the federal joint committee" in sample:
        return "Decisão (Beschluss)"
    if "tragende gründe zum beschluss" in sample or "supporting reasons for the decision" in sample:
        return "Razões Determinantes (Tragende Gründe)"
    if "dossier zur nutzenbewertung" in sample or "dossier for benefit assessment" in sample:
        return "Dossiê G-BA"

    return "Documento Geral G-BA"


def extract_active_substance(text: str, file_path: Path) -> str:
    """Localiza o princípio ativo (Wirkstoff / Active Substance)."""
    # 1. Tentar padrões regex específicos no texto
    patterns = [
        r"(?:Wirkstoff(?:e)?|Active substance(?:s)?|Active ingredient(?:s)?)\s*:\s*([^\n\r|;]{3,70})",
        r"(?:Wirkstoff(?:e)?|Active substance(?:s)?)\s*\n+\s*([A-Za-z0-9\-\s,]{3,60})",
        r"##+\s*(?:Wirkstoff|Active Substance)\s*[:\-]?\s*([^\n\r]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = clean_text_snippet(m.group(1), max_chars=80)
            if val and len(val) > 2 and not val.lower().startswith("tabelle") and not val.lower().startswith("table"):
                return val

    # 2. Tentar padrão de título em negrito no início do documento: **Nivolumab (Melanom...)**
    m_title = re.search(r"^\s*\*\*([A-Za-z0-9\-]+)(?:\s*\([^)]+\))?\*\*", text, re.MULTILINE)
    if m_title:
        val = m_title.group(1).strip()
        if len(val) > 2 and val.lower() not in ["tabelle", "table", "abbildung", "figure", "inhalt", "contents"]:
            return val.capitalize()

    # 3. Tentar deduzir do caminho da pasta (ex: active substance Nivolumab ou wirkstoff Nivolumab)
    path_str = str(file_path)
    m_folder = re.search(r"(?:active substance|wirkstoff)[_\s]+([A-Za-z0-9\-]+)", path_str, re.IGNORECASE)
    if m_folder:
        return m_folder.group(1).capitalize()

    # 4. Tentar deduzir do nome do arquivo
    m_name = re.search(r"verfahren[-_\s]+([A-Za-z0-9\-]+)", file_path.name, re.IGNORECASE)
    if m_name:
        return m_name.group(1).capitalize()

    return "Não identificado"


def extract_brand_name(text: str) -> str:
    """Extrai o nome comercial (Handelsname / Brand Name)."""
    patterns = [
        r"(?:Handelsname(?:n)?|Brand name(?:s)?|Trade name(?:s)?)\s*:\s*([^\n\r|;]{2,50})",
        r"(?:Handelsname|Trade name)\s*\n+\s*([A-Za-z0-9\-\s®™]{2,50})",
        r"\b([A-Z][a-z0-9\-]+[®™])\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = clean_text_snippet(m.group(1), max_chars=50)
            if val and len(val) > 2:
                return val
    return "Não identificado"


def extract_indication(text: str) -> str:
    """Localiza o trecho da indicação terapêutica (Anwendungsgebiet / Indication)."""
    patterns = [
        r"(?:Zugelassenes\s+)?Anwendungsgebiet(?:\s*\(gemäß\s+Zulassung\))?\s*[:\n]\s*([^\n\r#|]{15,400})",
        r"(?:Approved\s+)?Therapeutic\s+indication\s*[:\n]\s*([^\n\r#|]{15,400})",
        r"(?:Indikation|Indication)\s*:\s*([^\n\r#|]{15,400})",
        r"##\s*\*\*(?:Thema|Theme|Topic)\*\*\s*\n+\s*([^\n\r]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = clean_text_snippet(m.group(1), max_chars=350)
            if len(val) > 10:
                return val

    # Tentar extrair do título inicial em parênteses: **Nivolumab (Melanom, adjuvant...)**
    m_title = re.search(r"^\s*\*\*[A-Za-z0-9\-]+\s*\(([^)]+)\)\*\*", text, re.MULTILINE)
    if m_title:
        val = clean_text_snippet(m_title.group(1), max_chars=350)
        if len(val) > 5:
            return val

    return "Não especificada"


def extract_comparator_therapy(text: str) -> str:
    """Localiza a terapia comparadora apropriada (Zweckmäßige Vergleichstherapie / ZVT)."""
    patterns = [
        r"(?:Zweckmäßige\s+Vergleichstherapie|Appropriate\s+comparative\s+therapy|ZVT)\s*[:\n]\s*([^\n\r#|]{10,350})",
        r"(?:Vergleichstherapie|Comparator\s+therapy)\s*:\s*([^\n\r#|]{10,350})",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = clean_text_snippet(m.group(1), max_chars=300)
            if len(val) > 5:
                return val
    return "Não identificada"


def extract_additional_benefit(text: str) -> str:
    """Avalia o grau de benefício adicional concedido pelo G-BA."""
    text_lower = text.casefold()

    # Categorias oficiais do AMNOG / G-BA
    if "erheblicher zusatznutzen" in text_lower or "major additional benefit" in text_lower:
        return "Grande (Erheblich)"
    if "beträchtlicher zusatznutzen" in text_lower or "considerable additional benefit" in text_lower or "beträchtlich" in text_lower:
        return "Beträchtlich (Considerável)"
    if "geringer zusatznutzen" in text_lower or "minor additional benefit" in text_lower or "gering" in text_lower:
        return "Gering (Pequeno)"
    if "nicht quantifizierbarer zusatznutzen" in text_lower or "non-quantifiable additional benefit" in text_lower or "nicht quantifizierbar" in text_lower:
        return "Nicht quantifizierbar (Não Quantificável)"
    if "kein zusatznutzen belegt" in text_lower or "kein zusatznutzen" in text_lower or "no additional benefit" in text_lower:
        return "Kein Zusatznutzen (Sem Benefício Adicional)"
    if "geringerer nutzen" in text_lower or "lesser benefit" in text_lower:
        return "Geringerer Nutzen (Menor que o comparador)"
    if "nutzen nicht belegt" in text_lower or "benefit not proven" in text_lower:
        return "Nutzen nicht belegt (Não Comprovado)"

    return "Não mencionado / Em avaliação"


def extract_decision_date(text: str) -> str:
    """Extrai a data da decisão do G-BA."""
    patterns = [
        r"(?:Beschluss\s+vom|Decision\s+of)\s+([0-9]{1,2}\.\s+[A-Za-zäöü]+\s+[0-9]{4})",
        r"(?:Beschluss\s+vom|Decision\s+of|Stand:?)\s+([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{4})",
        r"\b([0-9]{4}-[0-9]{2}-[0-9]{2})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return "N/A"


def extract_orphan_status(text: str) -> str:
    """Verifica menção a medicamento órfão (Orphan Drug)."""
    text_lower = text.casefold()
    if "seltene leiden" in text_lower or "orphan drug" in text_lower or "orphan-arzneimittel" in text_lower:
        return "Sim (Orphan Drug)"
    return "Não / Padrão"


def parse_markdown_file(file_path: Path, relative_to: Path) -> Dict[str, Any]:
    """Lê um arquivo Markdown e extrai todas as informações clínicas e regulatórias."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            # Lemos os primeiros 30.000 caracteres (onde costumam estar sumários, dados da decisão e cabeçalhos)
            content = f.read(35000)
    except Exception as e:
        content = ""

    rel_path = str(file_path.relative_to(relative_to))
    procedure_id = extract_procedure_id(file_path)
    doc_type = identify_document_type(file_path, content)
    active_substance = extract_active_substance(content, file_path)
    brand_name = extract_brand_name(content)
    indication = extract_indication(content)
    comparator = extract_comparator_therapy(content)
    benefit = extract_additional_benefit(content)
    date = extract_decision_date(content)
    orphan = extract_orphan_status(content)

    return {
        "id_procedimento": procedure_id,
        "principio_ativo": active_substance,
        "nome_comercial": brand_name,
        "tipo_documento": doc_type,
        "beneficio_adicional": benefit,
        "status_orphan_drug": orphan,
        "indicacao_terapeutica": indication,
        "terapia_comparadora": comparator,
        "data_decisao": date,
        "arquivo_origem": file_path.name,
        "caminho_relativo": rel_path,
        "tamanho_kb": round(file_path.stat().st_size / 1024, 1),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Etapa 3: Extrator de Dados Regulatórios e Clínicos dos Markdowns do G-BA"
    )
    parser.add_argument(
        "--source",
        type=str,
        default=None,
        help="Pasta com arquivos Markdown (padrão: tenta 'gba_markdown_de', se não houver, tenta 'gba_markdown_en')",
    )
    parser.add_argument(
        "--output-prefix",
        type=str,
        default="dados_extraidos/dados_extraidos_gba",
        help="Prefixo dos arquivos de saída (gera .xlsx, .json e .csv) (padrão: dados_extraidos/dados_extraidos_gba)",
    )
    parser.add_argument("--limit", type=int, default=None, help="Limite máximo de arquivos para extrair")
    parser.add_argument("--filter", type=str, default=None, help="Filtro de texto no caminho (ex: '1219' ou 'Nivolumab')")

    args = parser.parse_args()

    # Seleciona a pasta de origem inteligentemente
    if args.source:
        source_dir = Path(args.source).resolve()
    else:
        de_dir = Path("gba_markdown_de").resolve()
        en_dir = Path("gba_markdown_en").resolve()
        if de_dir.exists() and any(de_dir.iterdir()):
            source_dir = de_dir
        elif en_dir.exists() and any(en_dir.iterdir()):
            source_dir = en_dir
        else:
            source_dir = de_dir

    if not source_dir.exists():
        console.print(f"[bold red]Erro:[/bold red] Pasta de origem não encontrada: {source_dir}")
        console.print("[yellow]Execute primeiro a Etapa 1: `python 1_converter_pdf_para_markdown.py`[/yellow]")
        sys.exit(1)

    console.print(
        Panel(
            f"[bold green]ETAPA 3: Extrator de Dados Regulatórios e Clínicos[/bold green]\n"
            f"[dim]Mineração de Princípio Ativo, Indicação, Terapia Comparadora e Zusatznutzen[/dim]\n\n"
            f"• [cyan]Origem:[/cyan] {source_dir}\n"
            f"• [cyan]Arquivos de saída:[/cyan] {args.output_prefix}.xlsx / .json / .csv\n"
            f"• [cyan]Filtro:[/cyan] {args.filter or 'Nenhum'}\n"
            f"• [cyan]Limite:[/cyan] {args.limit or 'Sem limite'}",
            title="[bold yellow]Pipeline G-BA[/bold yellow]",
            border_style="cyan",
        )
    )

    # Coletar arquivos .md
    console.print(f"[cyan][*][/cyan] Buscando arquivos Markdown em [yellow]{source_dir.name}[/yellow]...")
    md_files: List[Path] = []
    filter_norm = args.filter.casefold() if args.filter else None

    for root, _, files in os.walk(source_dir):
        for f in files:
            if f.lower().endswith(".md"):
                p = Path(root) / f
                if filter_norm and filter_norm not in str(p).casefold():
                    continue
                md_files.append(p)
                if args.limit and len(md_files) >= args.limit:
                    break
        if args.limit and len(md_files) >= args.limit:
            break

    total_found = len(md_files)
    console.print(f"[green][+][/green] Encontrados [bold]{total_found}[/bold] arquivos Markdown para análise.")

    if total_found == 0:
        console.print("[yellow]Nenhum arquivo markdown para extrair.[/yellow]")
        return

    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=35),
        MofNCompleteColumn(),
        TextColumn("[bold cyan]{task.percentage:>3.0f}%[/bold cyan]"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
    )

    overall_task = progress.add_task("[bold yellow]Extraindo Dados[/bold yellow]", total=len(md_files))

    extracted_records: List[Dict[str, Any]] = []
    t0 = time.time()

    with progress:
        for md_path in md_files:
            rec = parse_markdown_file(md_path, relative_to=source_dir)
            extracted_records.append(rec)
            progress.advance(overall_task, 1)

    elapsed = time.time() - t0

    # Converter para DataFrame Pandas
    df = pd.DataFrame(extracted_records)

    excel_out = Path(f"{args.output_prefix}.xlsx").resolve()
    json_out = Path(f"{args.output_prefix}.json").resolve()
    csv_out = Path(f"{args.output_prefix}.csv").resolve()
    excel_out.parent.mkdir(parents=True, exist_ok=True)

    console.print(f"\n[cyan][*][/cyan] Salvando dados extraídos...")

    # Salvar em Excel com formatação básica
    try:
        with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Dados Extraidos G-BA")
        console.print(f"[green]✔ Planilha Excel salva:[/green] [bold]{excel_out.name}[/bold]")
    except Exception as e:
        console.print(f"[red]Erro ao salvar Excel: {e}[/red]")

    # Salvar em JSON
    try:
        with open(json_out, "w", encoding="utf-8") as f_json:
            json.dump(extracted_records, f_json, ensure_ascii=False, indent=2)
        console.print(f"[green]✔ Arquivo JSON salvo:[/green] [bold]{json_out.name}[/bold]")
    except Exception as e:
        console.print(f"[red]Erro ao salvar JSON: {e}[/red]")

    # Salvar em CSV
    try:
        df.to_csv(csv_out, index=False, encoding="utf-8-sig")
        console.print(f"[green]✔ Arquivo CSV salvo:[/green] [bold]{csv_out.name}[/bold]")
    except Exception as e:
        console.print(f"[red]Erro ao salvar CSV: {e}[/red]")

    # Resumo no terminal
    table_summary = Table(title="Estatísticas da Extração", border_style="green")
    table_summary.add_column("Categoria", style="cyan")
    table_summary.add_column("Contagem", style="bold white", justify="right")

    table_summary.add_row("Total de documentos analisados", str(len(df)))
    table_summary.add_row("Procedimentos únicos identificados", str(df["id_procedimento"].nunique()))
    table_summary.add_row("Princípios ativos identificados", str(df[df["principio_ativo"] != "Não identificado"]["principio_ativo"].nunique()))
    table_summary.add_row("Tempo total de extração", f"{elapsed:.2f}s")

    console.print()
    console.print(table_summary)

    # Distribuição de Benefício Adicional
    if "beneficio_adicional" in df.columns:
        benefit_counts = df["beneficio_adicional"].value_counts()
        table_benefit = Table(title="Distribuição de Benefício Adicional (Zusatznutzen)", border_style="yellow")
        table_benefit.add_column("Grau de Benefício Adicional", style="magenta")
        table_benefit.add_column("Ocorrências", style="bold white", justify="right")
        for benefit_label, count in benefit_counts.items():
            table_benefit.add_row(str(benefit_label), str(count))
        console.print()
        console.print(table_benefit)

    console.print("\n[bold green]✔ Extração finalizada com sucesso![/bold green]\n")


if __name__ == "__main__":
    main()
