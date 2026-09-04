#!/usr/bin/env python3
"""
FORÇA TOTAL - Etapa 3: Extrator Regulatório Multi-Thread de Altíssima Velocidade
================================================================================
- Varredura concorrente com 8 threads simultâneas.
- Expressões regulares pré-compiladas em C para mineração em microssegundos.
- Exportação instantânea em lote para Excel (.xlsx), JSON (.json) e CSV (.csv).
- Geração de métricas completas no terminal.
"""

import argparse
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import psutil
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

# REGEX PRÉ-COMPILADAS PARA MÁXIMA PERFORMANCE
RE_PROCEDURE = re.compile(r"\b([A-Za-z]?-\d{3,5})\b")
RE_PROCEDURE_NUM = re.compile(r"\b(\d{3,5})\b")

RE_SUBSTANCE_TEXT = [
    re.compile(r"(?:Wirkstoff(?:e)?|Active substance(?:s)?|Active ingredient(?:s)?)\s*:\s*([^\n\r|;]{3,70})", re.IGNORECASE),
    re.compile(r"(?:Wirkstoff(?:e)?|Active substance(?:s)?)\s*\n+\s*([A-Za-z0-9\-\s,]{3,60})", re.IGNORECASE),
    re.compile(r"##+\s*(?:Wirkstoff|Active Substance)\s*[:\-]?\s*([^\n\r]+)", re.IGNORECASE),
]
RE_SUBSTANCE_TITLE = re.compile(r"^\s*\*\*([A-Za-z0-9\-]+)(?:\s*\([^)]+\))?\*\*", re.MULTILINE)
RE_SUBSTANCE_PATH = re.compile(r"(?:active substance|wirkstoff)[_\s]+([A-Za-z0-9\-]+)", re.IGNORECASE)
RE_SUBSTANCE_FILE = re.compile(r"verfahren[-_\s]+([A-Za-z0-9\-]+)", re.IGNORECASE)

RE_BRAND = [
    re.compile(r"(?:Handelsname(?:n)?|Brand name(?:s)?|Trade name(?:s)?)\s*:\s*([^\n\r|;]{2,50})", re.IGNORECASE),
    re.compile(r"(?:Handelsname|Trade name)\s*\n+\s*([A-Za-z0-9\-\s®™]{2,50})", re.IGNORECASE),
    re.compile(r"\b([A-Z][a-z0-9\-]+[®™])\b"),
]

RE_INDICATION = [
    re.compile(r"(?:Zugelassenes\s+)?Anwendungsgebiet(?:\s*\(gemäß\s+Zulassung\))?\s*[:\n]\s*([^\n\r#|]{15,400})", re.IGNORECASE),
    re.compile(r"(?:Approved\s+)?Therapeutic\s+indication\s*[:\n]\s*([^\n\r#|]{15,400})", re.IGNORECASE),
    re.compile(r"(?:Indikation|Indication)\s*:\s*([^\n\r#|]{15,400})", re.IGNORECASE),
    re.compile(r"##\s*\*\*(?:Thema|Theme|Topic)\*\*\s*\n+\s*([^\n\r]+)", re.IGNORECASE),
]
RE_INDICATION_TITLE = re.compile(r"^\s*\*\*[A-Za-z0-9\-]+\s*\(([^)]+)\)\*\*", re.MULTILINE)

RE_COMPARATOR = [
    re.compile(r"(?:Zweckmäßige\s+Vergleichstherapie|Appropriate\s+comparative\s+therapy|ZVT)\s*[:\n]\s*([^\n\r#|]{10,350})", re.IGNORECASE),
    re.compile(r"(?:Vergleichstherapie|Comparator\s+therapy)\s*:\s*([^\n\r#|]{10,350})", re.IGNORECASE),
]

RE_DATE = [
    re.compile(r"(?:Beschluss\s+vom|Decision\s+of)\s+([0-9]{1,2}\.\s+[A-Za-zäöü]+\s+[0-9]{4})", re.IGNORECASE),
    re.compile(r"(?:Beschluss\s+vom|Decision\s+of|Stand:?)\s+([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{4})", re.IGNORECASE),
    re.compile(r"\b([0-9]{4}-[0-9]{2}-[0-9]{2})\b"),
]


def find_project_root() -> Path:
    script_dir = Path(__file__).resolve().parent
    if (script_dir / "gba").exists() or (script_dir / "gba_markdown_de").exists():
        return script_dir
    if (script_dir.parent / "gba").exists() or (script_dir.parent / "gba_markdown_de").exists():
        return script_dir.parent
    return script_dir


def clean_snippet(text: str, max_chars: int = 400) -> str:
    if not text:
        return ""
    cleaned = re.sub(r"[\r\n\t]+", " ", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"[*#_`|]", "", cleaned).strip()
    return cleaned[:max_chars] if len(cleaned) <= max_chars else cleaned[: max_chars - 3] + "..."


def identify_document_type(file_path: Path, sample: str) -> str:
    p_str = str(file_path).casefold()
    s_str = sample.casefold()

    if "beschluss" in p_str or "decision" in p_str:
        return "Decisão (Beschluss)"
    if "tragende gruende" in p_str or "tragende gründe" in p_str or "supporting reasons" in p_str:
        return "Razões Determinantes (Tragende Gründe)"
    if "modul 1" in p_str or "module 1" in p_str:
        return "Dossiê - Módulo 1 (Geral)"
    if "modul 2" in p_str or "module 2" in p_str:
        return "Dossiê - Módulo 2 (Resumo Clínico)"
    if "modul 3" in p_str or "module 3" in p_str:
        return "Dossiê - Módulo 3 (Eficácia/Segurança)"
    if "modul 4" in p_str or "module 4" in p_str:
        return "Dossiê - Módulo 4 (Custos/Impacto)"
    if "modul 5" in p_str or "module 5" in p_str:
        return "Dossiê - Módulo 5 (Anexos)"
    if "iqwig" in p_str or "abschlussbericht" in p_str or "final report" in p_str:
        return "Avaliação IQWiG (Bericht)"
    if "stellungnahme" in p_str or "statement" in p_str:
        return "Manifestação (Stellungnahme)"

    if "beschluss des gemeinsamen bundesausschusses" in s_str or "decision of the federal joint committee" in s_str:
        return "Decisão (Beschluss)"
    if "tragende gründe zum beschluss" in s_str or "supporting reasons for the decision" in s_str:
        return "Razões Determinantes (Tragende Gründe)"
    if "dossier zur nutzenbewertung" in s_str or "dossier for benefit assessment" in s_str:
        return "Dossiê G-BA"

    return "Documento Geral G-BA"


def extract_additional_benefit(text_lower: str) -> str:
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


def parse_file_turbo(file_path: Path, source_dir: Path) -> Dict[str, Any]:
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read(35000)
    except Exception:
        content = ""

    content_lower = content.casefold()

    # ID do procedimento
    procedure_id = "N/A"
    for p in file_path.parts:
        m = RE_PROCEDURE.search(p)
        if m:
            procedure_id = m.group(1).upper()
            break
        m_num = RE_PROCEDURE_NUM.search(p)
        if m_num:
            procedure_id = f"D-{m_num.group(1)}"
            break

    # Princípio Ativo
    active_substance = "Não identificado"
    for pat in RE_SUBSTANCE_TEXT:
        m = pat.search(content)
        if m:
            val = clean_snippet(m.group(1), 80)
            if len(val) > 2 and not val.lower().startswith("tabelle") and not val.lower().startswith("table"):
                active_substance = val
                break

    if active_substance == "Não identificado":
        m_t = RE_SUBSTANCE_TITLE.search(content)
        if m_t:
            v = m_t.group(1).strip()
            if len(v) > 2 and v.lower() not in ["tabelle", "table", "abbildung", "figure", "inhalt", "contents"]:
                active_substance = v.capitalize()

    if active_substance == "Não identificado":
        m_f = RE_SUBSTANCE_PATH.search(str(file_path))
        if m_f:
            active_substance = m_f.group(1).capitalize()

    # Nome comercial
    brand = "Não identificado"
    for pat in RE_BRAND:
        m = pat.search(content)
        if m:
            val = clean_snippet(m.group(1), 50)
            if len(val) > 2:
                brand = val
                break

    # Indicação
    indication = "Não especificada"
    for pat in RE_INDICATION:
        m = pat.search(content)
        if m:
            val = clean_snippet(m.group(1), 350)
            if len(val) > 10:
                indication = val
                break

    if indication == "Não especificada":
        m_it = RE_INDICATION_TITLE.search(content)
        if m_it:
            v = clean_snippet(m_it.group(1), 350)
            if len(v) > 5:
                indication = v

    # Terapia Comparadora
    comparator = "Não identificada"
    for pat in RE_COMPARATOR:
        m = pat.search(content)
        if m:
            val = clean_snippet(m.group(1), 300)
            if len(val) > 5:
                comparator = val
                break

    # Data
    date_val = "N/A"
    for pat in RE_DATE:
        m = pat.search(content)
        if m:
            date_val = m.group(1).strip()
            break

    # Orphan Drug
    orphan = "Sim (Orphan Drug)" if ("seltene leiden" in content_lower or "orphan drug" in content_lower or "orphan-arzneimittel" in content_lower) else "Não / Padrão"

    return {
        "id_procedimento": procedure_id,
        "principio_ativo": active_substance,
        "nome_comercial": brand,
        "tipo_documento": identify_document_type(file_path, content),
        "beneficio_adicional": extract_additional_benefit(content_lower),
        "status_orphan_drug": orphan,
        "indicacao_terapeutica": indication,
        "terapia_comparadora": comparator,
        "data_decisao": date_val,
        "arquivo_origem": file_path.name,
        "caminho_relativo": str(file_path.relative_to(source_dir)),
        "tamanho_kb": round(file_path.stat().st_size / 1024, 1),
    }


def main():
    root = find_project_root()

    parser = argparse.ArgumentParser(
        description="FORÇA TOTAL: Extrator Regulatório e Clínico em Velocidade Máxima"
    )
    parser.add_argument(
        "--source",
        type=str,
        default=None,
        help="Pasta com arquivos Markdown (padrão: busca automática)",
    )
    parser.add_argument(
        "--output-prefix",
        type=str,
        default=str(root / "dados_extraidos" / "dados_extraidos_gba"),
    )
    parser.add_argument("--workers", type=int, default=8, help="Número de threads de extração (padrão: 8)")
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--filter", type=str, default=None)

    args = parser.parse_args()

    if args.source:
        source_dir = Path(args.source).resolve()
    else:
        de_dir = (root / "gba_markdown_de").resolve()
        en_dir = (root / "gba_markdown_en").resolve()
        source_dir = de_dir if (de_dir.exists() and any(de_dir.iterdir())) else en_dir

    if not source_dir.exists():
        console.print(f"[bold red]Erro:[/bold red] Pasta não encontrada: {source_dir}")
        sys.exit(1)

    workers = max(1, args.workers)

    console.print(
        Panel(
            f"[bold red]⚡ FORÇA TOTAL: Extrator Regulatório Multi-Thread ⚡[/bold red]\n"
            f"[dim]Varredura Paralela Instantânea com 8 Threads[/dim]\n\n"
            f"• [cyan]Origem:[/cyan] {source_dir}\n"
            f"• [cyan]Saídas:[/cyan] {args.output_prefix}.xlsx / .json / .csv\n"
            f"• [cyan]Threads de extração:[/cyan] [bold green]{workers} threads[/bold green]\n"
            f"• [cyan]Filtro:[/cyan] {args.filter or 'Todos'}\n"
            f"• [cyan]Limite:[/cyan] {args.limit or 'Sem limite'}",
            title="[bold yellow]Modo Força Total[/bold yellow]",
            border_style="red",
        )
    )

    console.print(f"[cyan][*][/cyan] Buscando Markdowns em [yellow]{source_dir.name}[/yellow]...")
    md_files: List[Path] = []
    filter_norm = args.filter.casefold() if args.filter else None

    for r, _, files in os.walk(source_dir):
        for f in files:
            if f.lower().endswith(".md"):
                p = Path(r) / f
                if filter_norm and filter_norm not in str(p).casefold():
                    continue
                md_files.append(p)
                if args.limit and len(md_files) >= args.limit:
                    break
        if args.limit and len(md_files) >= args.limit:
            break

    total_found = len(md_files)
    console.print(f"[green][+][/green] Total de Markdowns encontrados: [bold]{total_found}[/bold]")

    if total_found == 0:
        console.print("[yellow]Nenhum arquivo markdown para extrair.[/yellow]")
        return

    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=40),
        MofNCompleteColumn(),
        TextColumn("[bold cyan]{task.percentage:>3.0f}%[/bold cyan]"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
    )

    overall_task = progress.add_task("[bold red]Minerando Dados (Turbo)[/bold red]", total=len(md_files))

    extracted_records: List[Dict[str, Any]] = []
    t0 = time.time()

    with progress:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = [executor.submit(parse_file_turbo, p, source_dir) for p in md_files]
            for f in as_completed(futures):
                try:
                    rec = f.result()
                    extracted_records.append(rec)
                except Exception:
                    pass
                progress.advance(overall_task, 1)

    elapsed = max(0.01, time.time() - t0)

    df = pd.DataFrame(extracted_records)
    excel_out = Path(f"{args.output_prefix}.xlsx").resolve()
    json_out = Path(f"{args.output_prefix}.json").resolve()
    csv_out = Path(f"{args.output_prefix}.csv").resolve()
    excel_out.parent.mkdir(parents=True, exist_ok=True)

    console.print(f"\n[cyan][*][/cyan] Gravando saídas estruturadas...")

    try:
        with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Dados Extraidos G-BA")
        console.print(f"[green]✔ Planilha Excel salva:[/green] [bold]{excel_out}[/bold]")
    except Exception as e:
        console.print(f"[red]Erro ao salvar Excel: {e}[/red]")

    try:
        with open(json_out, "w", encoding="utf-8") as f_json:
            json.dump(extracted_records, f_json, ensure_ascii=False, indent=2)
        console.print(f"[green]✔ Arquivo JSON salvo:[/green] [bold]{json_out}[/bold]")
    except Exception as e:
        console.print(f"[red]Erro ao salvar JSON: {e}[/red]")

    try:
        df.to_csv(csv_out, index=False, encoding="utf-8-sig")
        console.print(f"[green]✔ Arquivo CSV salvo:[/green] [bold]{csv_out}[/bold]")
    except Exception as e:
        console.print(f"[red]Erro ao salvar CSV: {e}[/red]")

    table_summary = Table(title="Estatísticas Força Total (Etapa 3)", border_style="red")
    table_summary.add_column("Métrica", style="cyan")
    table_summary.add_column("Valor", style="bold white", justify="right")

    table_summary.add_row("Documentos minerados", str(len(df)))
    table_summary.add_row("Procedimentos únicos", str(df["id_procedimento"].nunique()))
    table_summary.add_row("Princípios ativos encontrados", str(df[df["principio_ativo"] != "Não identificado"]["principio_ativo"].nunique()))
    table_summary.add_row("Tempo total de extração", f"{elapsed:.2f}s")
    table_summary.add_row("Throughput", f"{(len(df) / elapsed):.1f} docs / segundo")

    console.print()
    console.print(table_summary)
    console.print(f"\n[bold green]✔ Extração Turbo finalizada![/bold green]\n")


if __name__ == "__main__":
    main()
