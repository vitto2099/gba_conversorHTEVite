#!/usr/bin/env python3
"""
Etapa 1: Conversor Rápido de PDF para Markdown (Original em Alemão)
==================================================================
- Converte PDFs da pasta G-BA para Markdown preservando tabelas, cabeçalhos
  e estrutura visual original com PyMuPDF4LLM.
- Converte no idioma ORIGINAL (Alemão), sem o atraso da tradução de texto inline.
- Organiza as pastas e nomes de arquivos de saída em INGLÊS (compatível com a planilha).
- Baixo impacto no PC: prioridade de processo reduzida (psutil) e controle térmico.
- Baixo consumo de RAM: conversão em lotes de páginas, sem vazamentos de memória.
- Interface visual no terminal com Rich: barra geral, progresso por páginas e status de CPU/RAM.
- Suporte a retoma automática (Resume): ignora arquivos já convertidos.
- Gravação atômica (.tmp -> .md) para segurança contra encerramentos abruptos.
"""

import argparse
import gc
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
import psutil
import pymupdf
import pymupdf4llm
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TaskID,
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

DEFAULT_NAME_TRANSLATIONS = {
    "nutzenbewertungsverfahren": "benefit assessment procedure",
    "nutzenbewertung": "benefit assessment",
    "verfahren": "procedure",
    "wirkstoff": "active substance",
    "wirkstoffe": "active substances",
    "erneute": "renewed",
    "neubewertung": "reassessment",
    "chronische": "chronic",
    "chronischer": "chronic",
    "chronisch": "chronic",
    "diabetes": "diabetes",
    "mellitus": "mellitus",
    "pulmonal": "pulmonary",
    "schizophrenie": "schizophrenia",
    "multiresistente": "multidrug resistant",
    "multiresistenter": "multidrug resistant",
    "störung": "disorder",
    "stoerung": "disorder",
    "überschuss": "excess",
    "ueberschuss": "excess",
    "überaktive": "overactive",
    "ueberaktive": "overactive",
    "hiv": "HIV",
    "infektion": "infection",
    "infektionen": "infections",
    "prostatakarzinom": "prostate cancer",
    "schilddruesenkarzinom": "thyroid cancer",
    "schilddrüsenkarzinom": "thyroid cancer",
    "mammakarzinom": "breast cancer",
    "lungenkarzinom": "lung cancer",
    "kolorektales": "colorectal",
    "karzinom": "carcinoma",
    "colitis": "colitis",
    "ulcerosa": "ulcerosa",
    "copd": "COPD",
    "neues": "new",
    "anwendungsgebiet": "therapeutic indication",
    "anwendungsgebiete": "therapeutic indications",
    "larven": "larvae",
    "lebende": "live",
    "von": "of",
    "aus": "from",
    "zum": "for the",
    "zur": "for the",
    "und": "and",
    "oder": "or",
    "mit": "with",
    "bei": "in",
    "dossier": "Dossier",
    "vergleichstherapie": "Comparative Therapy",
    "stellungnahmen": "Statements",
    "beschluesse": "Decisions",
    "beschluss": "Decision",
    "beschlusstext": "Decision text",
    "tragende": "supporting",
    "gruende": "reasons",
    "tragende gruende": "supporting reasons",
    "tragende gründe": "supporting reasons",
    "zusammenfassung": "Summary",
    "bericht": "Report",
    "abschlussbericht": "Final Report",
    "erwachsene": "adults",
    "kinder": "children",
    "jugendliche": "adolescents",
    "other": "Other",
    "english": "English",
}


def set_process_low_priority(eco_mode: bool = False):
    """Garante prioridade baixa no Windows para que o computador não trave nem congele."""
    if sys.platform == "win32":
        try:
            p = psutil.Process()
            if eco_mode:
                p.nice(psutil.IDLE_PRIORITY_CLASS)
            else:
                p.nice(psutil.BELOW_NORMAL_PRIORITY_CLASS)
        except Exception:
            pass


def normalize_key(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("_", " ").strip()).casefold()


def translate_name_part(name: str, translations: Dict[str, str]) -> str:
    stem = Path(name).stem
    suffix = Path(name).suffix
    norm = normalize_key(stem)

    if norm in translations:
        return translations[norm] + suffix

    tokens = re.split(r"([_\-\s(),]+)", stem)
    parts = []
    for token in tokens:
        k = normalize_key(token)
        if not k:
            parts.append(token)
            continue
        parts.append(translations.get(k, DEFAULT_NAME_TRANSLATIONS.get(k, token)))

    res = "".join(parts).replace("_", " ")
    res = re.sub(r"\s+", " ", res).strip()
    res = re.sub(r"\s+([),])", r"\1", res)
    res = re.sub(r"([(])\s+", r"\1", res)
    return res + suffix


def load_excel_mapping(excel_path: Path) -> Dict[str, str]:
    """Carrega o mapeamento de caminhos da planilha com cache JSON para carregamento instantâneo."""
    cache_path = excel_path.with_suffix(".json")
    if cache_path.exists():
        try:
            if cache_path.stat().st_mtime >= excel_path.stat().st_mtime:
                with open(cache_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass

    mapping: Dict[str, str] = {}
    if not excel_path.exists():
        return mapping

    console.print(f"[cyan][*][/cyan] Carregando mapeamento de pastas/arquivos: [yellow]{excel_path.name}[/yellow]...")
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    sheet = wb.active
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, None)

    orig_idx = 5
    en_idx = 6

    if header:
        for idx, col in enumerate(header):
            if col == "caminho_relativo_original":
                orig_idx = idx
            elif col == "caminho_relativo_ingles":
                en_idx = idx

    for r in rows:
        if r and len(r) > max(orig_idx, en_idx):
            orig = r[orig_idx]
            en = r[en_idx]
            if orig and en:
                mapping[str(orig).replace("/", "\\").casefold()] = str(en).replace("/", "\\")

    try:
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(mapping, f)
    except Exception:
        pass

    console.print(f"[green][+][/green] [bold]{len(mapping)}[/bold] caminhos mapeados via Excel.")
    return mapping


def get_target_path(pdf_path: Path, source_root: Path, output_root: Path, excel_map: Dict[str, str]) -> Path:
    rel_path = pdf_path.relative_to(source_root)
    rel_str_key = str(rel_path).casefold()

    if rel_str_key in excel_map:
        en_rel = Path(excel_map[rel_str_key])
    else:
        parts = [translate_name_part(p, DEFAULT_NAME_TRANSLATIONS) for p in rel_path.parts]
        en_rel = Path(*parts)

    target = output_root / en_rel
    return target.with_suffix(".md")


def convert_single_pdf_to_md(
    pdf_path: Path,
    target_md: Path,
    eco_mode: bool = False,
    page_batch_size: int = 25,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Tuple[bool, str, str, int]:
    """
    Converte um arquivo PDF para Markdown no idioma original em lotes de páginas,
    mantendo baixo consumo de memória RAM e gravando de forma atômica via .tmp.
    """
    temp_target = target_md.with_suffix(".tmp")
    target_md.parent.mkdir(parents=True, exist_ok=True)

    try:
        doc = pymupdf.open(str(pdf_path))
        total_pages = len(doc)

        if total_pages == 0:
            doc.close()
            return False, pdf_path.name, "PDF vazio", 0

        with open(temp_target, "w", encoding="utf-8") as f_out:
            for start_page in range(0, total_pages, page_batch_size):
                end_page = min(start_page + page_batch_size, total_pages)
                page_indices = list(range(start_page, end_page))

                if progress_callback:
                    progress_callback(start_page, total_pages, f"Pág {start_page + 1}-{end_page}/{total_pages}")

                # Extrai markdown para o lote de páginas preservando tabelas e títulos
                page_md = pymupdf4llm.to_markdown(doc, pages=page_indices)
                f_out.write(page_md)
                f_out.write("\n\n")
                f_out.flush()

                if progress_callback:
                    progress_callback(end_page, total_pages, f"Pág {end_page}/{total_pages} OK")

                if eco_mode:
                    time.sleep(0.02)

        doc.close()

        # Substituição atômica: renomeia .tmp para .md
        if temp_target.exists():
            if target_md.exists():
                target_md.unlink()
            temp_target.rename(target_md)

        size_kb = target_md.stat().st_size / 1024
        size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
        return True, pdf_path.name, size_str, total_pages

    except Exception as e:
        if temp_target.exists():
            try:
                temp_target.unlink()
            except Exception:
                pass
        return False, pdf_path.name, str(e), 0
    finally:
        gc.collect()


def get_hardware_status() -> Tuple[float, float, str]:
    """Retorna CPU%, RAM% e uso de RAM em GB para monitoramento."""
    try:
        cpu = psutil.cpu_percent(interval=None)
        mem = psutil.virtual_memory()
        ram_gb = mem.used / (1024**3)
        return cpu, mem.percent, f"{ram_gb:.1f} GB ({mem.percent:.0f}%)"
    except Exception:
        return 0.0, 0.0, "N/A"


def main():
    parser = argparse.ArgumentParser(
        description="Etapa 1: Conversão Rápida de PDF para Markdown no idioma Original (Alemão) com Pastas em Inglês"
    )
    parser.add_argument("--source", type=str, default="gba", help="Pasta com PDFs originais (padrão: gba)")
    parser.add_argument("--output", type=str, default="gba_markdown_de", help="Pasta destino (padrão: gba_markdown_de)")
    parser.add_argument(
        "--excel",
        type=str,
        default="nomes_pastas_arquivos_traduzidos.xlsx",
        help="Planilha de mapeamento de nomes traduzidos",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=2,
        help="Número de threads simultâneas (padrão: 2)",
    )
    parser.add_argument(
        "--eco",
        action="store_true",
        help="Modo Ultra-Leve: 1 thread, prioridade ociosa e pequenas pausas térmicas",
    )
    parser.add_argument("--limit", type=int, default=None, help="Limite máximo de arquivos para converter")
    parser.add_argument("--filter", type=str, default=None, help="Filtro de texto no caminho (ex: '1219' ou 'Nivolumab')")

    args = parser.parse_args()

    # Configuração de prioridade do processo
    set_process_low_priority(eco_mode=args.eco)

    source_dir = Path(args.source).resolve()
    output_dir = Path(args.output).resolve()
    excel_path = Path(args.excel).resolve()

    if not source_dir.exists():
        console.print(f"[bold red]Erro:[/bold red] Pasta de origem não encontrada: {source_dir}")
        sys.exit(1)

    workers = 1 if args.eco else max(1, args.workers)

    # Banner Inicial
    console.print(
        Panel(
            f"[bold green]ETAPA 1: Conversor de PDF para Markdown (Alemão Original)[/bold green]\n"
            f"[dim]Pastas de saída em Inglês | Tabelas preservadas | Baixo consumo de CPU/RAM[/dim]\n\n"
            f"• [cyan]Origem:[/cyan] {source_dir}\n"
            f"• [cyan]Destino:[/cyan] {output_dir}\n"
            f"• [cyan]Threads:[/cyan] {workers} {'(Modo Eco ativado)' if args.eco else ''}\n"
            f"• [cyan]Filtro:[/cyan] {args.filter or 'Nenhum (todos os arquivos)'}\n"
            f"• [cyan]Limite:[/cyan] {args.limit or 'Sem limite'}",
            title="[bold yellow]Pipeline G-BA[/bold yellow]",
            border_style="cyan",
        )
    )

    excel_map = load_excel_mapping(excel_path)

    # Coleta de arquivos PDF
    console.print(f"[cyan][*][/cyan] Escaneando PDFs em [yellow]{source_dir.name}[/yellow]...")
    pdf_files: List[Path] = []
    filter_norm = args.filter.casefold() if args.filter else None

    for root, _, files in os.walk(source_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                p = Path(root) / f
                if filter_norm:
                    if filter_norm not in str(p).casefold():
                        continue
                pdf_files.append(p)
                if args.limit and len(pdf_files) >= args.limit:
                    break
        if args.limit and len(pdf_files) >= args.limit:
            break

    total_found = len(pdf_files)
    console.print(f"[green][+][/green] Encontrados [bold]{total_found}[/bold] arquivos PDF para análise.")

    if total_found == 0:
        console.print("[yellow]Nenhum arquivo para processar. Finalizando.[/yellow]")
        return

    # Verificar quais já foram convertidos (Resume)
    tasks_to_run: List[Tuple[Path, Path]] = []
    skipped_count = 0

    for pdf in pdf_files:
        target = get_target_path(pdf, source_dir, output_dir, excel_map)
        if target.exists() and target.stat().st_size > 0:
            skipped_count += 1
        else:
            tasks_to_run.append((pdf, target))

    console.print(
        f"[cyan][*][/cyan] Já convertidos anteriormente: [bold]{skipped_count}[/bold] | "
        f"Restantes para converter: [bold]{len(tasks_to_run)}[/bold]"
    )

    if not tasks_to_run:
        console.print("[bold green]✔ Todos os arquivos já estão convertidos em Markdown![/bold green]")
        return

    # Configuração de barras de progresso Rich
    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=35),
        MofNCompleteColumn(),
        TextColumn("[bold cyan]{task.percentage:>3.0f}%[/bold cyan]"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
        refresh_per_second=4,
    )

    overall_task = progress.add_task("[bold yellow]Progresso Geral[/bold yellow]", total=len(tasks_to_run))
    file_task = progress.add_task("[dim cyan]Arquivo Atual[/dim cyan]", total=100)

    sucessos = 0
    erros = 0
    total_paginas = 0
    t0 = time.time()

    def update_file_progress(current: int, total: int, msg: str):
        if total > 0:
            progress.update(
                file_task,
                completed=int((current / total) * 100),
                description=f"[cyan]{msg}[/cyan]",
            )

    with progress:
        if workers == 1:
            # Processamento sequencial (mais suave para RAM e temperatura)
            for pdf_path, target_md in tasks_to_run:
                short_name = pdf_path.name[:32] + "..." if len(pdf_path.name) > 35 else pdf_path.name
                progress.update(file_task, completed=0, description=f"[cyan]{short_name}[/cyan]")

                cpu, mem_pct, ram_str = get_hardware_status()
                progress.update(
                    overall_task,
                    description=f"[bold yellow]Geral[/bold yellow] [dim](RAM: {ram_str} | CPU: {cpu:.0f}%)[/dim]",
                )

                ok, name, msg, pgs = convert_single_pdf_to_md(
                    pdf_path=pdf_path,
                    target_md=target_md,
                    eco_mode=args.eco,
                    page_batch_size=25,
                    progress_callback=update_file_progress,
                )

                if ok:
                    sucessos += 1
                    total_paginas += pgs
                else:
                    erros += 1

                progress.advance(overall_task, 1)
        else:
            # Processamento com pool de threads
            def worker_fn(item: Tuple[Path, Path]):
                p_in, p_out = item
                return convert_single_pdf_to_md(
                    pdf_path=p_in,
                    target_md=p_out,
                    eco_mode=args.eco,
                    page_batch_size=25,
                    progress_callback=None,
                )

            with ThreadPoolExecutor(max_workers=workers) as executor:
                for ok, name, msg, pgs in executor.map(worker_fn, tasks_to_run):
                    if ok:
                        sucessos += 1
                        total_paginas += pgs
                    else:
                        erros += 1

                    cpu, mem_pct, ram_str = get_hardware_status()
                    progress.update(
                        overall_task,
                        description=f"[bold yellow]Geral[/bold yellow] [dim](RAM: {ram_str} | CPU: {cpu:.0f}%)[/dim]",
                    )
                    progress.advance(overall_task, 1)

    elapsed = time.time() - t0
    elapsed_str = f"{elapsed:.1f}s" if elapsed < 60 else f"{elapsed/60:.1f} min"

    summary_table = Table(title="Resumo da Conversão (Etapa 1)", border_style="green")
    summary_table.add_column("Métrica", style="cyan", justify="left")
    summary_table.add_column("Valor", style="bold white", justify="right")

    summary_table.add_row("Arquivos convertidos com sucesso", str(sucessos))
    summary_table.add_row("Arquivos já existentes (pulados)", str(skipped_count))
    summary_table.add_row("Falhas", f"[red]{erros}[/red]" if erros > 0 else "0")
    summary_table.add_row("Total de páginas processadas", str(total_paginas))
    summary_table.add_row("Tempo total de execução", elapsed_str)
    if sucessos > 0:
        summary_table.add_row("Média por arquivo", f"{(elapsed / sucessos):.2f}s")

    console.print()
    console.print(summary_table)
    console.print(f"[bold green]✔ Concluído com sucesso![/bold green] Arquivos salvos em: [cyan]{output_dir}[/cyan]\n")


if __name__ == "__main__":
    main()
