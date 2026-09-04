#!/usr/bin/env python3
"""
FORÇA TOTAL - Etapa 2: Tradutor de Markdown Paralelo de Alta Velocidade
=====================================================================
- Tradução paralela multi-thread (processa múltiplos arquivos simultaneamente).
- Sem pausas artificiais desnecessárias.
- Backoff inteligente individual: se houver rate limit (429), apenas o worker afetado pausa.
- Resume instantâneo: ignora arquivos já traduzidos.
- Monitoramento de arquivos/minuto em tempo real via Rich.
"""

import argparse
import gc
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Optional, Tuple

import psutil
from deep_translator import GoogleTranslator
from deep_translator.exceptions import TranslationNotFound
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


def find_project_root() -> Path:
    script_dir = Path(__file__).resolve().parent
    if (script_dir / "gba").exists() or (script_dir / "gba_markdown_de").exists():
        return script_dir
    if (script_dir.parent / "gba").exists() or (script_dir.parent / "gba_markdown_de").exists():
        return script_dir.parent
    return script_dir


def set_turbo_priority():
    if sys.platform == "win32":
        try:
            p = psutil.Process()
            p.nice(psutil.ABOVE_NORMAL_PRIORITY_CLASS)
        except Exception:
            pass


def sanitize_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\x00", "").replace("\ufffd", " ")
    return text.replace("\r\n", "\n").replace("\r", "\n")


class TurboTranslator:
    def __init__(self, source_lang: str = "de", target_lang: str = "en"):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.translator = GoogleTranslator(source=self.source_lang, target=self.target_lang)

    def _chunk_text(self, text: str, max_chunk_size: int = 3400) -> List[str]:
        paragraphs = text.split("\n\n")
        chunks = []
        current: List[str] = []
        current_len = 0

        for p in paragraphs:
            p_len = len(p)
            if current_len + p_len + 2 > max_chunk_size:
                if current:
                    chunks.append("\n\n".join(current))
                    current = [p]
                    current_len = p_len
                else:
                    lines = p.split("\n")
                    sub_chunk: List[str] = []
                    sub_len = 0
                    for line in lines:
                        if sub_len + len(line) + 1 > max_chunk_size:
                            if sub_chunk:
                                chunks.append("\n".join(sub_chunk))
                                sub_chunk = [line]
                                sub_len = len(line)
                            else:
                                chunks.append(line[:max_chunk_size])
                                sub_chunk = [line[max_chunk_size:]]
                                sub_len = len(sub_chunk[0])
                        else:
                            sub_chunk.append(line)
                            sub_len += len(line) + 1
                    if sub_chunk:
                        chunks.append("\n".join(sub_chunk))
                    current = []
                    current_len = 0
            else:
                current.append(p)
                current_len += p_len + 2

        if current:
            chunks.append("\n\n".join(current))

        return chunks

    def translate_chunk(self, chunk: str) -> str:
        clean = sanitize_text(chunk)
        if not clean.strip():
            return chunk

        if clean.strip().startswith("```") and clean.strip().endswith("```"):
            return chunk

        for attempt in range(3):
            try:
                res = self.translator.translate(clean)
                time.sleep(0.01)  # latência mínima
                return res if res else chunk
            except TranslationNotFound:
                return chunk
            except Exception as e:
                err_str = str(e).lower()
                if "429" in err_str or "500" in err_str or "connection" in err_str:
                    time.sleep(1.0 * (attempt + 1))
                else:
                    time.sleep(0.2 * (attempt + 1))
                if attempt == 2:
                    return chunk
        return chunk

    def translate_document(self, content: str) -> str:
        if not content or not content.strip():
            return content
        chunks = self._chunk_text(content)
        translated = [self.translate_chunk(c) for c in chunks]
        return "\n\n".join(translated)


def translate_file_task(source_file: Path, target_file: Path) -> Tuple[bool, str]:
    temp_target = target_file.with_suffix(".tmp")
    target_file.parent.mkdir(parents=True, exist_ok=True)

    try:
        with open(source_file, "r", encoding="utf-8", errors="replace") as f:
            content = f.read()

        translator = TurboTranslator(source_lang="de", target_lang="en")
        translated_text = translator.translate_document(content)

        with open(temp_target, "w", encoding="utf-8", errors="replace") as f_out:
            f_out.write(translated_text)

        if temp_target.exists():
            if target_file.exists():
                target_file.unlink()
            temp_target.rename(target_file)

        return True, "OK"
    except Exception as e:
        if temp_target.exists():
            try:
                temp_target.unlink()
            except Exception:
                pass
        return False, str(e)


def main():
    root = find_project_root()

    parser = argparse.ArgumentParser(
        description="FORÇA TOTAL: Tradutor Paralelo de Markdown em Velocidade Máxima"
    )
    parser.add_argument("--source", type=str, default=str(root / "gba_markdown_de"))
    parser.add_argument("--output", type=str, default=str(root / "gba_markdown_en"))
    parser.add_argument(
        "--workers",
        type=int,
        default=4,
        help="Número de threads simultâneas de tradução (padrão: 4 para alta velocidade sem bloqueio)",
    )
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--filter", type=str, default=None)

    args = parser.parse_args()

    set_turbo_priority()

    source_dir = Path(args.source).resolve()
    output_dir = Path(args.output).resolve()

    if not source_dir.exists():
        console.print(f"[bold red]Erro:[/bold red] Pasta de origem não encontrada: {source_dir}")
        sys.exit(1)

    workers = max(1, args.workers)

    console.print(
        Panel(
            f"[bold red]⚡ FORÇA TOTAL: Tradutor Multi-Thread de Markdowns ⚡[/bold red]\n"
            f"[dim]Tradução Concorrente de Arquivos | Backoff Individual | Turbo ativado[/dim]\n\n"
            f"• [cyan]Origem:[/cyan] {source_dir}\n"
            f"• [cyan]Destino:[/cyan] {output_dir}\n"
            f"• [cyan]Trabalhadores paralelos:[/cyan] [bold green]{workers} threads[/bold green]\n"
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
    console.print(f"[green][+][/green] Total de Markdowns localizados: [bold]{total_found}[/bold]")

    tasks_to_run: List[Tuple[Path, Path]] = []
    skipped = 0

    for md_path in md_files:
        rel = md_path.relative_to(source_dir)
        target = output_dir / rel
        if target.exists() and target.stat().st_size > 0:
            skipped += 1
        else:
            tasks_to_run.append((md_path, target))

    console.print(
        f"[cyan][*][/cyan] Já traduzidos: [bold]{skipped}[/bold] | "
        f"Restantes para traduzir: [bold]{len(tasks_to_run)}[/bold]"
    )

    if not tasks_to_run:
        console.print("[bold green]✔ Todos os arquivos já estão traduzidos![/bold green]")
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
        refresh_per_second=4,
    )

    overall_task = progress.add_task("[bold red]Traduzindo em Paralelo[/bold red]", total=len(tasks_to_run))

    sucessos = 0
    erros = 0
    t0 = time.time()

    with progress:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = [executor.submit(translate_file_task, src, tgt) for src, tgt in tasks_to_run]
            for f in as_completed(futures):
                try:
                    ok, _ = f.result()
                    if ok:
                        sucessos += 1
                    else:
                        erros += 1
                except Exception:
                    erros += 1
                progress.advance(overall_task, 1)

    total_time = max(0.1, time.time() - t0)

    table = Table(title="Desempenho Força Total (Etapa 2)", border_style="red")
    table.add_column("Métrica", style="cyan")
    table.add_column("Resultado", style="bold white", justify="right")

    table.add_row("Arquivos traduzidos", str(sucessos))
    table.add_row("Erros", f"[red]{erros}[/red]" if erros > 0 else "0")
    table.add_row("Tempo total", f"{total_time:.1f}s")
    table.add_row("Velocidade média", f"{(sucessos / total_time):.2f} arq/s")

    console.print()
    console.print(table)
    console.print(f"[bold green]✔ Tradução concluída![/bold green] Salvos em: [cyan]{output_dir}[/cyan]\n")


if __name__ == "__main__":
    main()
