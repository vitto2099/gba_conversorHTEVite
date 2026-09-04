#!/usr/bin/env python3
"""
FORÇA TOTAL - Etapa 1: Conversor PDF para Markdown de Altíssima Velocidade
========================================================================
- Modo Turbo com múltiplos workers paralelos (padrão: 8 threads).
- Sem pausas artificiais (latência zero).
- Prioridade de CPU elevada (Normal / Acima do Normal) para usar o potencial da máquina.
- Conversão direta em memória para máximo throughput com PyMuPDF4LLM.
- Gravação atômica (.tmp -> .md) com proteção contra arquivos corrompidos.
- Pula automaticamente arquivos já convertidos (Resume instantâneo).
- Monitoramento de velocidade (arquivos/segundo) em tempo real via Rich.
"""

import argparse
import gc
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Tuple

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


def find_project_root() -> Path:
    """Localiza a pasta raiz do projeto automaticamente."""
    script_dir = Path(__file__).resolve().parent
    if (script_dir / "gba").exists():
        return script_dir
    if (script_dir.parent / "gba").exists():
        return script_dir.parent
    return script_dir


def set_turbo_priority():
    """Configura prioridade normal/alta para liberar potência de CPU."""
    if sys.platform == "win32":
        try:
            p = psutil.Process()
            p.nice(psutil.ABOVE_NORMAL_PRIORITY_CLASS)
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
    """Carrega o mapeamento com cache JSON ultra veloz."""
    cache_path = excel_path.with_suffix(".json")
    if cache_path.exists():
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    mapping: Dict[str, str] = {}
    if not excel_path.exists():
        return mapping

    console.print(f"[cyan][*][/cyan] Carregando mapa de nomes: [yellow]{excel_path.name}[/yellow]...")
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


def convert_single_pdf_turbo(pdf_path: Path, target_md: Path) -> Tuple[bool, str, int]:
    """Conversão ultra-rápida sem pausas artificiais."""
    temp_target = target_md.with_suffix(".tmp")
    target_md.parent.mkdir(parents=True, exist_ok=True)

    try:
        doc = pymupdf.open(str(pdf_path))
        total_pages = len(doc)

        if total_pages == 0:
            doc.close()
            return False, "PDF vazio", 0

        # Para PDFs menores de 80 páginas, converte direto na memória (muito mais rápido)
        if total_pages <= 80:
            md_content = pymupdf4llm.to_markdown(doc)
            doc.close()
            with open(temp_target, "w", encoding="utf-8", errors="replace") as f_out:
                f_out.write(md_content)
        else:
            # Para PDFs gigantes (>80 páginas), converte em blocos de 50 páginas para balancear CPU e RAM
            with open(temp_target, "w", encoding="utf-8", errors="replace") as f_out:
                for start_page in range(0, total_pages, 50):
                    end_page = min(start_page + 50, total_pages)
                    page_indices = list(range(start_page, end_page))
                    page_md = pymupdf4llm.to_markdown(doc, pages=page_indices)
                    f_out.write(page_md)
                    f_out.write("\n\n")
            doc.close()

        # Substituição atômica
        if temp_target.exists():
            if target_md.exists():
                target_md.unlink()
            temp_target.rename(target_md)

        return True, "OK", total_pages

    except Exception as e:
        if temp_target.exists():
            try:
                temp_target.unlink()
            except Exception:
                pass
        return False, str(e), 0


def main():
    root = find_project_root()

    parser = argparse.ArgumentParser(
        description="FORÇA TOTAL: Conversor de PDF para Markdown em Velocidade Máxima"
    )
    parser.add_argument("--source", type=str, default=str(root / "gba"), help="Pasta de origem dos PDFs")
    parser.add_argument("--output", type=str, default=str(root / "gba_markdown_de"), help="Pasta de saída dos Markdowns")
    parser.add_argument("--excel", type=str, default=str(root / "nomes_pastas_arquivos_traduzidos.xlsx"))
    parser.add_argument(
        "--workers",
        type=int,
        default=8,
        help="Número de threads simultâneas (padrão: 8 para Core i5 13ª geração)",
    )
    parser.add_argument("--limit", type=int, default=None, help="Limite máximo de arquivos para converter")
    parser.add_argument("--filter", type=str, default=None, help="Filtro no caminho (ex: '1219')")

    args = parser.parse_args()

    set_turbo_priority()

    source_dir = Path(args.source).resolve()
    output_dir = Path(args.output).resolve()
    excel_path = Path(args.excel).resolve()

    if not source_dir.exists():
        console.print(f"[bold red]Erro:[/bold red] Pasta de origem não encontrada: {source_dir}")
        sys.exit(1)

    workers = max(1, args.workers)

    console.print(
        Panel(
            f"[bold red]⚡ FORÇA TOTAL: Conversor Ultra-Rápido de PDF para Markdown ⚡[/bold red]\n"
            f"[dim]Paralelismo Máximo | Latência Zero | Prioridade Alta de CPU[/dim]\n\n"
            f"• [cyan]Origem:[/cyan] {source_dir}\n"
            f"• [cyan]Destino:[/cyan] {output_dir}\n"
            f"• [cyan]Threads paralelas:[/cyan] [bold green]{workers} threads ativas[/bold green]\n"
            f"• [cyan]Filtro:[/cyan] {args.filter or 'Todos'}\n"
            f"• [cyan]Limite:[/cyan] {args.limit or 'Sem limite'}",
            title="[bold yellow]Modo Força Total[/bold yellow]",
            border_style="red",
        )
    )

    excel_map = load_excel_mapping(excel_path)

    console.print(f"[cyan][*][/cyan] Varrendo PDFs em [yellow]{source_dir.name}[/yellow]...")
    pdf_files: List[Path] = []
    filter_norm = args.filter.casefold() if args.filter else None

    for r, _, files in os.walk(source_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                p = Path(r) / f
                if filter_norm and filter_norm not in str(p).casefold():
                    continue
                pdf_files.append(p)
                if args.limit and len(pdf_files) >= args.limit:
                    break
        if args.limit and len(pdf_files) >= args.limit:
            break

    total_found = len(pdf_files)
    console.print(f"[green][+][/green] Total de PDFs localizados: [bold]{total_found}[/bold]")

    tasks_to_run: List[Tuple[Path, Path]] = []
    skipped_count = 0

    for pdf in pdf_files:
        target = get_target_path(pdf, source_dir, output_dir, excel_map)
        if target.exists() and target.stat().st_size > 0:
            skipped_count += 1
        else:
            tasks_to_run.append((pdf, target))

    console.print(
        f"[cyan][*][/cyan] Já convertidos: [bold]{skipped_count}[/bold] | "
        f"Restantes para converter: [bold]{len(tasks_to_run)}[/bold]"
    )

    if not tasks_to_run:
        console.print("[bold green]✔ Todos os arquivos já estão convertidos![/bold green]")
        return

    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=40),
        MofNCompleteColumn(),
        TextColumn("[bold cyan]{task.percentage:>3.0f}%[/bold cyan]"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        TextColumn("[bold yellow]{task.fields[speed]}[/bold yellow]"),
        console=console,
        refresh_per_second=6,
    )

    overall_task = progress.add_task(
        "[bold red]Processando (Turbo)[/bold red]",
        total=len(tasks_to_run),
        speed="0.0 arq/s",
    )

    sucessos = 0
    erros = 0
    total_paginas = 0
    t0 = time.time()

    def process_item(item: Tuple[Path, Path]):
        p_in, p_out = item
        ok, msg, pgs = convert_single_pdf_turbo(p_in, p_out)
        return ok, pgs

    with progress:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = [executor.submit(process_item, item) for item in tasks_to_run]

            for future in as_completed(futures):
                try:
                    ok, pgs = future.result()
                    if ok:
                        sucessos += 1
                        total_paginas += pgs
                    else:
                        erros += 1
                except Exception:
                    erros += 1

                elapsed_now = max(0.1, time.time() - t0)
                current_speed = (sucessos + erros) / elapsed_now
                speed_str = f"{current_speed:.1f} arq/s"

                progress.update(overall_task, advance=1, speed=speed_str)

    total_time = max(0.1, time.time() - t0)
    final_speed = sucessos / total_time

    table = Table(title="Desempenho Força Total (Etapa 1)", border_style="red")
    table.add_column("Métrica", style="cyan")
    table.add_column("Resultado", style="bold white", justify="right")

    table.add_row("Arquivos convertidos", str(sucessos))
    table.add_row("Páginas processadas", str(total_paginas))
    table.add_row("Erros", f"[red]{erros}[/red]" if erros > 0 else "0")
    table.add_row("Tempo total", f"{total_time:.1f}s")
    table.add_row("Velocidade média", f"{final_speed:.2f} arquivos / segundo")

    console.print()
    console.print(table)
    console.print(f"[bold green]✔ Finalizado![/bold green] Markdowns gravados em: [cyan]{output_dir}[/cyan]\n")


if __name__ == "__main__":
    main()
