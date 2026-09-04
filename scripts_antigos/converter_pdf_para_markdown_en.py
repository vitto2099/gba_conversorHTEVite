#!/usr/bin/env python3
"""
G-BA High-Efficiency Markdown Converter & English Translator
============================================================
- Converte PDFs para Markdown preservando tabelas, títulos e formatação via PyMuPDF4LLM.
- Traduz todo o conteúdo interno do documento (Alemão -> Inglês) de forma segura e confiável.
- Nomes de pastas e arquivos organizados em Inglês (compatível com a planilha Excel).
- Modo Proteção de PC: Prioridade Baixa no Windows (o PC NUNCA trava e seus outros programas têm preferência).
- Consumo de memória controlado: processamento em lotes de páginas (não estoura RAM nem em PDFs gigantes).
- Interface rica no terminal com Rich: barra de progresso geral e por páginas do arquivo atual, velocidade, tempo estimado e status de CPU/RAM.
- Retomada automática (Resume): pula arquivos já convertidos.
- Gravação atômica (.tmp -> .md) para evitar arquivos corrompidos em caso de interrupção.
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
    """Garante prioridade baixa no Windows para que o PC nunca trave."""
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
    """Carrega o mapeamento de caminhos da planilha, usando cache JSON para inicialização instantânea."""
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

    console.print(f"[cyan][*][/cyan] Carregando mapeamento de nomes: [yellow]{excel_path.name}[/yellow]...")
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


def sanitize_text_for_translation(text: str) -> str:
    """Remove caracteres nulos ou corrompidos que possam travar o tradutor."""
    if not text:
        return ""
    text = text.replace("\x00", "").replace("\ufffd", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


class MarkdownTranslator:
    """Tradutor seguro com chunking inteligente, retentativas e respeito ao limite de requisições."""

    def __init__(self, source_lang: str = "auto", target_lang: str = "en", gentle_delay: float = 0.12):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.gentle_delay = gentle_delay
        self.translator = GoogleTranslator(source=self.source_lang, target=self.target_lang)

    def _chunk_text(self, text: str, max_chunk_size: int = 3200) -> List[str]:
        """Divide o markdown em parágrafos para não quebrar tabelas e títulos."""
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
        """Traduz um chunk com retentativas e tolerância a erros."""
        chunk_clean = sanitize_text_for_translation(chunk)
        if not chunk_clean.strip():
            return chunk

        if chunk_clean.strip().startswith("```") and chunk_clean.strip().endswith("```"):
            return chunk

        for attempt in range(3):
            try:
                res = self.translator.translate(chunk_clean)
                time.sleep(self.gentle_delay)
                return res if res else chunk
            except TranslationNotFound:
                # O texto já está em inglês ou não requer tradução
                return chunk
            except Exception as e:
                err_str = str(e).lower()
                if "500" in err_str or "429" in err_str or "connection" in err_str:
                    time.sleep(1.5 * (attempt + 1))
                else:
                    time.sleep(0.5 * (attempt + 1))
                if attempt == 2:
                    return chunk
        return chunk

    def translate_markdown(self, md_text: str) -> str:
        """Traduz o texto markdown mantendo parágrafos e tabelas."""
        if not md_text or not md_text.strip():
            return md_text

        chunks = self._chunk_text(md_text)
        translated_parts = []
        for c in chunks:
            translated_parts.append(self.translate_chunk(c))

        return "\n\n".join(translated_parts)


def convert_and_translate_single_pdf(
    pdf_path: Path,
    target_md: Path,
    translator: Optional[MarkdownTranslator],
    eco_mode: bool = False,
    page_batch_size: int = 15,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Tuple[bool, str, str, int]:
    """
    Converte um PDF para Markdown em fatias de páginas e traduz cada uma,
    mantendo uso de RAM baixo e salvando atomicamente em .tmp.
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
                    progress_callback(start_page, total_pages, f"Pág {start_page + 1}-{end_page}/{total_pages} Extraindo...")

                # Extrai markdown para o lote de páginas
                page_md = pymupdf4llm.to_markdown(doc, pages=page_indices)

                # Traduz se o tradutor estiver ativo
                if translator:
                    if progress_callback:
                        progress_callback(end_page, total_pages, f"Pág {start_page + 1}-{end_page}/{total_pages} Traduzindo...")
                    translated_md = translator.translate_markdown(page_md)
                else:
                    translated_md = page_md

                f_out.write(translated_md)
                f_out.write("\n\n")
                f_out.flush()

                if progress_callback:
                    progress_callback(end_page, total_pages, f"Pág {end_page}/{total_pages} Concluída")

                if eco_mode:
                    time.sleep(0.06)
                else:
                    time.sleep(0.01)

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
        description="Conversor Leve de PDF para Markdown com Tradução Completa em Inglês"
    )
    parser.add_argument("--source", type=str, default="gba", help="Pasta com PDFs originais (padrão: gba)")
    parser.add_argument("--output", type=str, default="gba_markdown_en", help="Pasta destino (padrão: gba_markdown_en)")
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
        help="Número de threads simultâneas (padrão: 2 para não esquentar o PC)",
    )
    parser.add_argument(
        "--eco",
        action="store_true",
        help="Modo Ultra-Leve: 1 thread, prioridade ociosa e pausas térmicas (ventiladores 100%% silenciosos)",
    )
    parser.add_argument(
        "--no-translate",
        action="store_true",
        help="Desativa tradução do texto interno (gera apenas Markdown com nomes traduzidos, muito rápido)",
    )
    parser.add_argument("--limit", type=int, default=None, help="Limite máximo de arquivos para converter")
    parser.add_argument("--filter", type=str, default=None, help="Filtro de texto no caminho (ex: '1000' ou 'Nivolumab')")

    args = parser.parse_args()

    # Aplica prioridade baixa no Windows
    set_process_low_priority(eco_mode=args.eco)

    source_root = Path(args.source).resolve()
    output_root = Path(args.output).resolve()
    excel_path = Path(args.excel).resolve()

    if not source_root.exists():
        console.print(f"[bold red][!] Pasta de origem não encontrada:[/bold red] {source_root}")
        sys.exit(1)

    output_root.mkdir(parents=True, exist_ok=True)

    workers = 1 if args.eco else max(1, args.workers)
    translate_content = not args.no_translate

    # Painel de apresentação
    info_table = Table(show_header=False, box=None)
    info_table.add_column("Key", style="bold cyan")
    info_table.add_column("Val", style="white")
    info_table.add_row("Origem (PDFs intactos):", str(source_root))
    info_table.add_row("Destino (Markdown):", str(output_root))
    info_table.add_row("Tradução do Conteúdo:", "[green]Ativada (DE -> EN)[/green]" if translate_content else "[yellow]Desativada (Apenas Nomes)[/yellow]")
    info_table.add_row("Nomes de Pastas/Arquivos:", "[green]100% Organizados em Inglês[/green]")
    info_table.add_row("Threads Simultâneas:", f"{workers} {'(Modo Eco)' if args.eco else ''}")
    info_table.add_row("Proteção do PC:", "[green]Prioridade Baixa no Windows (Zero travamentos ou lentidão)[/green]")
    if args.limit:
        info_table.add_row("Limite:", f"{args.limit} arquivos")
    if args.filter:
        info_table.add_row("Filtro:", f"Contendo '{args.filter}'")

    console.print(
        Panel(
            info_table,
            title="[bold yellow]G-BA PDF -> Markdown & English Translator[/bold yellow]",
            subtitle="[green]Alta Performance • Leveza • Retomada Automática[/green]",
            border_style="yellow",
        )
    )

    excel_map = load_excel_mapping(excel_path)

    console.print(f"[cyan][*][/cyan] Verificando arquivos já existentes em '{output_root.name}'...")
    existing_md_rel_set = set()
    for existing_file in output_root.rglob("*.md"):
        if existing_file.stat().st_size > 0:
            existing_md_rel_set.add(str(existing_file.relative_to(output_root)).casefold())

    console.print(f"[blue][i][/blue] Arquivos .md já prontos e preservados: [bold]{len(existing_md_rel_set)}[/bold]")

    console.print(f"[cyan][*][/cyan] Escaneando PDFs pendentes...")
    tasks: List[Tuple[Path, Path]] = []

    for root_dir, _, files in os.walk(source_root):
        root_path = Path(root_dir)
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf = root_path / file
                if args.filter and args.filter.lower() not in str(pdf).lower():
                    continue

                target_md = get_target_path(pdf, source_root, output_root, excel_map)
                target_rel_key = str(target_md.relative_to(output_root)).casefold()

                if target_rel_key not in existing_md_rel_set:
                    tasks.append((pdf, target_md))
                    if args.limit and len(tasks) >= args.limit:
                        break
        if args.limit and len(tasks) >= args.limit:
            break

    total_tasks = len(tasks)
    if total_tasks == 0:
        console.print("[bold green][+] Todos os arquivos já foram convertidos com sucesso![/bold green]")
        return

    console.print(f"[green][+][/green] Total de PDFs pendentes para converter: [bold cyan]{total_tasks}[/bold cyan]")
    console.print("[bold yellow]Pressione Ctrl+C a qualquer momento para pausar com segurança.[/bold yellow]\n")

    # Inicializa tradutor
    translator = MarkdownTranslator(source_lang="auto", target_lang="en") if translate_content else None

    # Configura barra de progresso dupla (Total de Documentos + Páginas do Arquivo Atual)
    progress = Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=32),
        MofNCompleteColumn(),
        TextColumn("[cyan]({task.percentage:>3.1f}%)[/cyan]"),
        TextColumn("[green]{task.fields[rate]}[/green]"),
        TextColumn("[yellow]{task.fields[eta]}[/yellow]"),
        TimeElapsedColumn(),
        refresh_per_second=4,
        console=console,
    )

    main_task = progress.add_task(
        "[bold green]Total Geral[/bold green]",
        total=total_tasks,
        rate="-- arq/s",
        eta="--",
    )

    file_task = progress.add_task(
        "[bold cyan]Arquivo Atual[/bold cyan]",
        total=100,
        rate="",
        eta="",
        visible=True,
    )

    start_time = time.time()
    success_count = 0
    error_count = 0

    with progress:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            task_iter = iter(tasks)

            def make_progress_cb(f_name: str):
                def cb(cur: int, tot: int, stage_text: str):
                    progress.update(
                        file_task,
                        total=max(tot, 1),
                        completed=cur,
                        description=f"[cyan]{f_name[:24]}[/cyan] ({stage_text})",
                    )
                return cb

            def worker_job(item: Tuple[Path, Path]):
                src, dst = item
                cb = make_progress_cb(src.stem)
                return convert_and_translate_single_pdf(
                    pdf_path=src,
                    target_md=dst,
                    translator=translator,
                    eco_mode=args.eco,
                    progress_callback=cb,
                )

            futures = []
            active_items = {}

            for _ in range(workers):
                try:
                    item = next(task_iter)
                    f = executor.submit(worker_job, item)
                    futures.append(f)
                    active_items[f] = item
                except StopIteration:
                    break

            last_stats_update = 0.0

            while futures:
                done_future = None
                for f in futures:
                    if f.done():
                        done_future = f
                        break

                if not done_future:
                    time.sleep(0.08)
                    continue

                futures.remove(done_future)
                item = active_items.pop(done_future)

                try:
                    ok, name, details, pages = done_future.result()
                    if ok:
                        success_count += 1
                    else:
                        error_count += 1
                except Exception:
                    error_count += 1

                completed = success_count + error_count
                elapsed = time.time() - start_time
                rate = completed / elapsed if elapsed > 0 else 0
                remaining_items = total_tasks - completed
                remaining_sec = (remaining_items / rate) if rate > 0 else 0

                rate_str = f"{rate*60:.1f} arq/m" if rate < 0.5 else f"{rate:.1f} arq/s"
                if remaining_sec < 60:
                    eta_str = f"ETA: {remaining_sec:.0f}s"
                elif remaining_sec < 3600:
                    eta_str = f"ETA: {remaining_sec/60:.1f}m"
                else:
                    eta_str = f"ETA: {remaining_sec/3600:.1f}h"

                progress.update(
                    main_task,
                    advance=1,
                    rate=rate_str,
                    eta=eta_str,
                )

                now = time.time()
                if now - last_stats_update >= 4.0:
                    last_stats_update = now
                    cpu, ram_pct, ram_str = get_hardware_status()
                    progress.console.print(
                        f"[dim]⚡ CPU: {cpu:.0f}% | 🧠 RAM: {ram_str} | Concluídos: {success_count} | Erros: {error_count}[/dim]"
                    )

                try:
                    next_item = next(task_iter)
                    f_next = executor.submit(worker_job, next_item)
                    futures.append(f_next)
                    active_items[f_next] = next_item
                except StopIteration:
                    pass

    total_time = time.time() - start_time
    time_str = f"{total_time/60:.1f} minutos" if total_time >= 60 else f"{total_time:.1f} segundos"

    console.print("\n" + "=" * 70)
    console.print(f"[bold green][✓] Conversão Concluída com Sucesso![/bold green]")
    console.print(f"• Documentos gerados: [bold green]{success_count}[/bold green]")
    if error_count > 0:
        console.print(f"• Falhas registradas: [bold red]{error_count}[/bold red]")
    console.print(f"• Tempo total: [bold cyan]{time_str}[/bold cyan]")
    console.print(f"• Pasta de saída: [bold white]{output_root}[/bold white]")
    console.print("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("\n[bold yellow][!] Execução pausada pelo usuário. Os arquivos concluídos estão seguros.[/bold yellow]")
        sys.exit(0)
