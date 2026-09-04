#!/usr/bin/env python3
"""
G-BA Concurrent PDF Downloader (Processador dos IDs Restantes)
"""

import csv
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from threading import Lock
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_URL = "https://www.g-ba.de"
PROCEDURE_URL = "https://www.g-ba.de/bewertungsverfahren/nutzenbewertung/{id}/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
}

SECTION_MAP = {
    "dossier": "01_Dossier",
    "zweckmaessige-vergleichstherapie": "02_Vergleichstherapie",
    "nutzenbewertung": "03_Nutzenbewertung",
    "stellungnahmen": "04_Stellungnahmen",
    "beschluesse": "05_Beschluesse",
    "english": "06_English",
}
DEFAULT_SECTION = "00_Other"

csv_lock = Lock()
print_lock = Lock()


def sanitize(text: str) -> str:
    text = text.strip()
    text = re.sub(r"[^\w\s\-_\(\)]", "_", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text)
    return text[:120]


def extract_drug_name(soup) -> str:
    h1 = soup.find("h1")
    if h1:
        return sanitize(h1.get_text())
    title = soup.find("title")
    if title:
        return sanitize(title.get_text().split("|")[0])
    return "Unknown"


def fetch_pdf_links(procedure_id: int, session: requests.Session):
    url = PROCEDURE_URL.format(id=procedure_id)
    try:
        resp = session.get(url, timeout=30)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
    except requests.RequestException as e:
        with print_lock:
            print(f"  [!] ID {procedure_id} erro ao buscar página: {e}", flush=True)
        return None

    soup = BeautifulSoup(resp.text, "html.parser")
    drug_name = extract_drug_name(soup)

    current_section = DEFAULT_SECTION
    section_pdfs = {}

    for tag in soup.find_all(True):
        tag_id = tag.get("id", "")
        if tag_id in SECTION_MAP:
            current_section = SECTION_MAP[tag_id]

        if tag.name == "a":
            href = tag.get("href", "")
            if href.lower().endswith(".pdf"):
                full_url = urljoin(BASE_URL, href)
                raw_name = tag.get_text(strip=True) or Path(urlparse(href).path).name
                raw_name = sanitize(raw_name)
                if not raw_name.lower().endswith(".pdf"):
                    raw_name += ".pdf"
                section_pdfs.setdefault(current_section, [])
                entry = (full_url, raw_name)
                if entry not in section_pdfs[current_section]:
                    section_pdfs[current_section].append(entry)

    return drug_name, section_pdfs


def download_pdf(url: str, dest: Path, session: requests.Session) -> bool:
    for attempt in range(3):
        try:
            resp = session.get(url, timeout=60, stream=True)
            resp.raise_for_status()
            dest.parent.mkdir(parents=True, exist_ok=True)
            temp_dest = dest.with_suffix(dest.suffix + ".tmp")
            with open(temp_dest, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    f.write(chunk)
            if temp_dest.exists() and temp_dest.stat().st_size > 0:
                temp_dest.replace(dest)
                return True
        except Exception:
            if attempt < 2:
                time.sleep(1.0)
            else:
                return False
    return False


def process_single_procedure(procedure_id: int, output_dir: Path, manifest_path: Path):
    session = requests.Session()
    session.headers.update(HEADERS)

    result = fetch_pdf_links(procedure_id, session)
    if not result:
        return {"id": procedure_id, "status": "not_found", "downloaded": 0, "skipped": 0, "total": 0}

    drug_name, section_pdfs = result
    proc_folder = output_dir / f"{procedure_id}_{drug_name[:60]}"
    total_found = sum(len(v) for v in section_pdfs.values())

    downloaded = 0
    skipped = 0
    errors = 0
    log_rows = []

    for section, pdfs in section_pdfs.items():
        for pdf_url, filename in pdfs:
            dest = proc_folder / section / filename
            status = ""

            if dest.exists() and dest.stat().st_size > 0:
                skipped += 1
                status = "skipped"
            else:
                ok = download_pdf(pdf_url, dest, session)
                if ok:
                    downloaded += 1
                    status = "downloaded"
                else:
                    errors += 1
                    status = "error"
                time.sleep(0.05)

            log_rows.append({
                "procedure_id": procedure_id,
                "drug": drug_name,
                "section": section,
                "filename": filename,
                "url": pdf_url,
                "local_path": str(dest),
                "status": status,
            })

    # Thread-safely append to manifest
    if log_rows and manifest_path:
        with csv_lock:
            file_exists = manifest_path.exists()
            with open(manifest_path, "a", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=["procedure_id", "drug", "section", "filename", "url", "local_path", "status"],
                )
                if not file_exists or manifest_path.stat().st_size == 0:
                    writer.writeheader()
                writer.writerows(log_rows)

    return {
        "id": procedure_id,
        "drug": drug_name,
        "status": "ok",
        "downloaded": downloaded,
        "skipped": skipped,
        "errors": errors,
        "total": total_found,
    }


def main():
    output_dir = Path("gba_pdfs")
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = output_dir / "download_manifest.csv"

    ids_file = Path("ids_498_em_diante.txt")
    if not ids_file.exists():
        print("Arquivo ids_498_em_diante.txt não encontrado.")
        return

    with open(ids_file, "r", encoding="utf-8") as f:
        target_ids = [int(line.strip()) for line in f if line.strip().isdigit()]

    total_procs = len(target_ids)
    print(f"=== Iniciando Download Concorrente dos Procedimentos (498 em diante) ===")
    print(f"Total de procedimentos na fila: {total_procs}")
    print(f"Faixa de IDs: {min(target_ids)} até {max(target_ids)}")
    print(f"Pasta de saída: {output_dir.resolve()}")
    print(f"Workers: 8 conexões concorrentes\n", flush=True)

    completed_count = 0
    total_dl = 0
    total_sk = 0
    total_err = 0
    start_time = time.time()

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {
            executor.submit(process_single_procedure, pid, output_dir, manifest_path): pid
            for pid in target_ids
        }

        for future in as_completed(futures):
            pid = futures[future]
            try:
                res = future.result()
                completed_count += 1
                total_dl += res.get("downloaded", 0)
                total_sk += res.get("skipped", 0)
                total_err += res.get("errors", 0)

                elapsed = time.time() - start_time
                avg_time = elapsed / completed_count
                remaining_time = avg_time * (total_procs - completed_count)
                rem_min = int(remaining_time // 60)
                rem_sec = int(remaining_time % 60)

                drug_short = res.get("drug", "")[:40]
                dl = res.get("downloaded", 0)
                sk = res.get("skipped", 0)
                tot = res.get("total", 0)

                with print_lock:
                    print(
                        f"[{completed_count}/{total_procs}] ({completed_count*100/total_procs:.1f}%) "
                        f"ID {res['id']:<4} | PDFs: {tot} (baixados: {dl}, pulados: {sk}) | "
                        f"{drug_short} | Restam aprox: {rem_min}m{rem_sec:02d}s",
                        flush=True,
                    )
            except Exception as e:
                with print_lock:
                    print(f"[!] ID {pid} falhou com erro: {e}", flush=True)

    print("\n" + "=" * 60)
    print("DOWNLOAD CONCLUÍDO COM SUCESSO!")
    print(f"Procedimentos processados: {completed_count}")
    print(f"Total PDFs baixados nesta sessão: {total_dl}")
    print(f"Total PDFs já existentes (pulados): {total_sk}")
    print(f"Erros: {total_err}")
    print(f"Tempo total: {int((time.time() - start_time) // 60)} min")
    print(f"Arquivos salvos em: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
