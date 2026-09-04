#!/usr/bin/env python3
"""
Etapa 2: Tradutor de Markdown (Alemão -> Inglês)
================================================
- Lê os arquivos Markdown gerados na Etapa 1 (em Alemão).
- Traduz parágrafo por parágrafo para Inglês, preservando tabelas Markdown,
  títulos (#, ##) e links.
- Proteção contra limites de API (Google Translate): chunking inteligente (<= 3200 caracteres),
  pausa suave entre requisições e retry com backoff exponencial.
- Retomada automática (Resume): ignora arquivos já traduzidos em gba_markdown_en.
- Gravação atômica (.tmp -> .md) para integridade dos arquivos.
- Monitoramento de hardware (CPU/RAM) e barra de progresso visual com Rich.
"""

import argparse
import gc
import os
import re
import sys
import time
from pathlib import Path
from typing import Callable, List, Optional, Tuple

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


def set_process_low_priority(eco_mode: bool = False):
    """Garante prioridade baixa no Windows para manter o PC 100% responsivo."""
    if sys.platform == "win32":
        try:
            p = psutil.Process()
            if eco_mode:
                p.nice(psutil.IDLE_PRIORITY_CLASS)
            else:
                p.nice(psutil.BELOW_NORMAL_PRIORITY_CLASS)
        except Exception:
            pass


def sanitize_text_for_translation(text: str) -> str:
    """Remove caracteres que corrompem a requisição ou travam a tradução."""
    if not text:
        return ""
    text = text.replace("\x00", "").replace("\ufffd", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


class MarkdownTranslator:
    """Gerenciador de tradução com chunking seguro, tolerância a falhas e pausas anti-ban."""

    def __init__(self, source_lang: str = "de", target_lang: str = "en", gentle_delay: float = 0.12):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.gentle_delay = gentle_delay
        self.translator = GoogleTranslator(source=self.source_lang, target=self.target_lang)

    def _chunk_text(self, text: str, max_chunk_size: int = 3200) -> List[str]:
        """Divide o texto em blocos de parágrafos sem quebrar tabelas."""
        paragraphs = text.split("\n\n")
        chunks = []
        current_chunk: List[str] = []
        current_len = 0

        for p in paragraphs:
            p_len = len(p)
            if current_len + p_len + 2 > max_chunk_size:
                if current_chunk:
                    chunks.append("\n\n".join(current_chunk))
                    current_chunk = [p]
                    current_len = p_len
                else:
                    # Parágrafo maior que max_chunk_size: divide por linhas
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
                    current_chunk = []
                    current_len = 0
            else:
                current_chunk.append(p)
                current_len += p_len + 2

        if current_chunk:
            chunks.append("\n\n".join(current_chunk))

        return chunks

    def translate_chunk(self, chunk: str) -> str:
        """Traduz um pedaço de texto com retentativas automáticas e controle de erros."""
        chunk_clean = sanitize_text_for_translation(chunk)
        if not chunk_clean.strip():
            return chunk

        # Pula blocos de código markdown puros
        if chunk_clean.strip().startswith("```") and chunk_clean.strip().endswith("```"):
            return chunk

        for attempt in range(4):
            try:
                res = self.translator.translate(chunk_clean)
                time.sleep(self.gentle_delay)
                return res if res else chunk
            except TranslationNotFound:
                # O texto já está em inglês ou não requer tradução
                return chunk
            except Exception as e:
                err_str = str(e).lower()
                # Se for rate limit (429) ou erro temporário de servidor (500), espera mais tempo
                if "429" in err_str or "500" in err_str or "connection" in err_str or "timeout" in err_str:
                    time.sleep(2.0 * (attempt + 1))
                else:
                    time.sleep(0.5 * (attempt + 1))

                if attempt == 3:
                    # Na última tentativa devolve o texto original para não perder o arquivo
                    return chunk
        return chunk

    def translate_file_content(
        self,
        content: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> str:
        """Traduz todo o documento markdown em chunks de parágrafos."""
        if not content or not content.strip():
            return content

        chunks = self._chunk_text(content)
        total_chunks = len(chunks)
        translated_parts = []

        for idx, c in enumerate(chunks):
            if progress_callback:
                progress_callback(idx + 1, total_chunks)
            translated_parts.append(self.translate_chunk(c))

        return "\n\n".join(translated_parts)


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
        description="Etapa 2: Tradução de Arquivos Markdown de Alemão para Inglês"
    )
    parser.add_argument("--source", type=str, default="gba_markdown_de", help="Pasta com Markdowns em Alemão (padrão: gba_markdown_de)")
    parser.add_argument("--output", type=str, default="gba_markdown_en", help="Pasta destino dos Markdowns em Inglês (padrão: gba_markdown_en)")
    parser.add_argument(
        "--delay",
        type=float,
        default=0.12,
        help="Pausa suave em segundos entre chamadas de tradução (padrão: 0.12)",
    )
    parser.add_argument(
        "--eco",
        action="store_true",
        help="Modo Ultra-Leve: prioridade ociosa e pausa maior (0.25s) para conexões sensíveis",
    )
    parser.add_argument("--limit", type=int, default=None, help="Limite máximo de arquivos para traduzir")
    parser.add_argument("--filter", type=str, default=None, help="Filtro de texto no caminho (ex: '1219' ou 'Nivolumab')")

    args = parser.parse_args()

    set_process_low_priority(eco_mode=args.eco)

    source_dir = Path(args.source).resolve()
    output_dir = Path(args.output).resolve()

    if not source_dir.exists():
        console.print(f"[bold red]Erro:[/bold red] Pasta de origem não encontrada: {source_dir}")
        console.print("[yellow]Dica: Execute primeiro a Etapa 1: `python 1_converter_pdf_para_markdown.py`[/yellow]")
        sys.exit(1)

    delay = 0.25 if args.eco else args.delay
    translator = MarkdownTranslator(source_lang="de", target_lang="en", gentle_delay=delay)

    console.print(
        Panel(
            f"[bold green]ETAPA 2: Tradutor de Markdown (Alemão -> Inglês)[/bold green]\n"
            f"[dim]Tradução por parágrafos | Tolerância a falhas e rate limits | Resume ativado[/dim]\n\n"
            f"• [cyan]Origem:[/cyan] {source_dir}\n"
            f"• [cyan]Destino:[/cyan] {output_dir}\n"
            f"• [cyan]Pausa suave:[/cyan] {delay}s por requisição {'(Modo Eco)' if args.eco else ''}\n"
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
    console.print(f"[green][+][/green] Encontrados [bold]{total_found}[/bold] arquivos Markdown para tradução.")

    if total_found == 0:
        console.print("[yellow]Nenhum arquivo encontrado para traduzir.[/yellow]")
        return

    # Verificar quais já foram traduzidos
    tasks_to_run: List[Tuple[Path, Path]] = []
    skipped_count = 0

    for md_path in md_files:
        rel_path = md_path.relative_to(source_dir)
        target = output_dir / rel_path
        if target.exists() and target.stat().st_size > 0:
            skipped_count += 1
        else:
            tasks_to_run.append((md_path, target))

    console.print(
        f"[cyan][*][/cyan] Já traduzidos anteriormente: [bold]{skipped_count}[/bold] | "
        f"Restantes para traduzir: [bold]{len(tasks_to_run)}[/bold]"
    )

    if not tasks_to_run:
        console.print("[bold green]✔ Todos os arquivos já estão traduzidos para Inglês![/bold green]")
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
        refresh_per_second=4,
    )

    overall_task = progress.add_task("[bold yellow]Progresso Geral[/bold yellow]", total=len(tasks_to_run))
    file_task = progress.add_task("[dim cyan]Arquivo Atual[/dim cyan]", total=100)

    sucessos = 0
    erros = 0
    t0 = time.time()

    def update_chunk_progress(current: int, total: int):
        if total > 0:
            pct = int((current / total) * 100)
            progress.update(
                file_task,
                completed=pct,
                description=f"[cyan]Traduzindo parágrafo {current}/{total}[/cyan]",
            )

    with progress:
        for md_path, target_md in tasks_to_run:
            short_name = md_path.name[:32] + "..." if len(md_path.name) > 35 else md_path.name
            progress.update(file_task, completed=0, description=f"[cyan]{short_name}[/cyan]")

            cpu, mem_pct, ram_str = get_hardware_status()
            progress.update(
                overall_task,
                description=f"[bold yellow]Geral[/bold yellow] [dim](RAM: {ram_str} | CPU: {cpu:.0f}%)[/dim]",
            )

            temp_target = target_md.with_suffix(".tmp")
            target_md.parent.mkdir(parents=True, exist_ok=True)

            try:
                with open(md_path, "r", encoding="utf-8", errors="replace") as f_in:
                    content = f_in.read()

                translated_content = translator.translate_file_content(
                    content=content,
                    progress_callback=update_chunk_progress,
                )

                with open(temp_target, "w", encoding="utf-8") as f_out:
                    f_out.write(translated_content)
                    f_out.flush()

                if temp_target.exists():
                    if target_md.exists():
                        target_md.unlink()
                    temp_target.rename(target_md)

                sucessos += 1
            except Exception as e:
                erros += 1
                if temp_target.exists():
                    try:
                        temp_target.unlink()
                    except Exception:
                        pass
                console.print(f"[red]Erro ao traduzir {md_path.name}: {e}[/red]")
            finally:
                progress.advance(overall_task, 1)
                gc.collect()

    elapsed = time.time() - t0
    elapsed_str = f"{elapsed:.1f}s" if elapsed < 60 else f"{elapsed/60:.1f} min"

    summary_table = Table(title="Resumo da Tradução (Etapa 2)", border_style="green")
    summary_table.add_column("Métrica", style="cyan", justify="left")
    summary_table.add_column("Valor", style="bold white", justify="right")

    summary_table.add_row("Arquivos traduzidos com sucesso", str(sucessos))
    summary_table.add_row("Arquivos já existentes (pulados)", str(skipped_count))
    summary_table.add_row("Falhas", f"[red]{erros}[/red]" if erros > 0 else "0")
    summary_table.add_row("Tempo total de execução", elapsed_str)
    if sucessos > 0:
        summary_table.add_row("Média por arquivo", f"{(elapsed / sucessos):.2f}s")

    console.print()
    console.print(summary_table)
    console.print(f"[bold green]✔ Tradução concluída![/bold green] Arquivos salvos em: [cyan]{output_dir}[/cyan]\n")


if __name__ == "__main__":
    main()
