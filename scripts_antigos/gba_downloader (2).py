#!/usr/bin/env python3
"""
G-BA Nutzenbewertung PDF Downloader
====================================
Downloads all PDFs from G-BA drug assessment pages.

Usage:
    # Single procedure by ID:
    python gba_downloader.py --ids 1102 1292

    # Range of IDs:
    python gba_downloader.py --range 1100 1300

    # From a text file (one ID per line):
    python gba_downloader.py --file ids.txt

    # Output directory (default: ./gba_pdfs):
    python gba_downloader.py --ids 1102 1292 --output ./my_folder

    # Dry run (list PDFs without downloading):
    python gba_downloader.py --ids 1102 --dry-run

    # Skip existing files (resume interrupted downloads):
    python gba_downloader.py --ids 1102 1292 --skip-existing
"""

import argparse
import csv
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Missing dependencies. Install them with:")
    print("    pip install requests beautifulsoup4")
    sys.exit(1)

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

# Section anchor → human-readable folder name
SECTION_MAP = {
    "dossier": "01_Dossier",
    "zweckmaessige-vergleichstherapie": "02_Vergleichstherapie",
    "nutzenbewertung": "03_Nutzenbewertung",
    "stellungnahmen": "04_Stellungnahmen",
    "beschluesse": "05_Beschluesse",
    "english": "06_English",
}
DEFAULT_SECTION = "00_Other"


def sanitize(text: str) -> str:
    """Make a string safe for use as a filename/folder name."""
    text = text.strip()
    text = re.sub(r"[^\w\s\-_\(\)]", "_", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text)
    return text[:120]


def get_section_for_link(link_tag, soup) -> str:
    """Walk DOM backwards to find which <h2>/<section> contains this link."""
    # Try to find the nearest preceding h2 or section with an id
    for ancestor in link_tag.parents:
        anchor_id = ancestor.get("id", "")
        if anchor_id in SECTION_MAP:
            return SECTION_MAP[anchor_id]
        # Check sibling h2 headings above
        for sibling in ancestor.find_all_previous(["h2", "section", "div"]):
            sid = sibling.get("id", "")
            if sid in SECTION_MAP:
                return SECTION_MAP[sid]
            break
    return DEFAULT_SECTION


def extract_drug_name(soup) -> str:
    """Extract the drug/procedure name from the page title."""
    h1 = soup.find("h1")
    if h1:
        return sanitize(h1.get_text())
    title = soup.find("title")
    if title:
        return sanitize(title.get_text().split("|")[0])
    return "Unknown"


def fetch_pdf_links(procedure_id: int, session: requests.Session) -> dict:
    """
    Fetch the procedure page and return a dict:
        { section_folder: [ (pdf_url, filename), ... ] }
    """
    url = PROCEDURE_URL.format(id=procedure_id)
    try:
        resp = session.get(url, timeout=30)
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"  [!] Could not fetch {url}: {e}")
        return {}

    soup = BeautifulSoup(resp.text, "html.parser")
    drug_name = extract_drug_name(soup)

    # Build a map: anchor_id → section label by scanning the page sequentially
    # We assign each <a href="*.pdf"> to its nearest preceding heading with a known id
    current_section = DEFAULT_SECTION
    section_pdfs: dict[str, list] = {}

    # Walk all elements in document order
    for tag in soup.find_all(True):
        tag_id = tag.get("id", "")
        if tag_id in SECTION_MAP:
            current_section = SECTION_MAP[tag_id]

        if tag.name == "a":
            href = tag.get("href", "")
            if href.lower().endswith(".pdf"):
                full_url = urljoin(BASE_URL, href)
                # Derive a clean filename
                raw_name = tag.get_text(strip=True) or Path(urlparse(href).path).name
                raw_name = sanitize(raw_name)
                if not raw_name.lower().endswith(".pdf"):
                    raw_name += ".pdf"
                # Deduplicate within section
                section_pdfs.setdefault(current_section, [])
                entry = (full_url, raw_name)
                if entry not in section_pdfs[current_section]:
                    section_pdfs[current_section].append(entry)

    return drug_name, section_pdfs


def download_pdf(url: str, dest: Path, session: requests.Session) -> bool:
    """Download a single PDF to dest. Returns True on success."""
    try:
        resp = session.get(url, timeout=60, stream=True)
        resp.raise_for_status()
        dest.parent.mkdir(parents=True, exist_ok=True)
        with open(dest, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)
        return True
    except requests.RequestException as e:
        print(f"      [!] Download failed: {e}")
        return False


def process_procedure(
    procedure_id: int,
    output_dir: Path,
    session: requests.Session,
    dry_run: bool = False,
    skip_existing: bool = True,
    log_rows: list = None,
) -> dict:
    """Process one procedure ID. Returns stats dict."""
    print(f"\n{'='*60}")
    print(f"  Procedure ID: {procedure_id}")
    url = PROCEDURE_URL.format(id=procedure_id)
    print(f"  URL: {url}")

    result = fetch_pdf_links(procedure_id, session)
    if not result:
        return {"id": procedure_id, "status": "fetch_error", "downloaded": 0, "skipped": 0}

    drug_name, section_pdfs = result
    print(f"  Drug/Title: {drug_name}")

    # Procedure folder: output_dir / {id}_{drug_name}
    proc_folder = output_dir / f"{procedure_id}_{drug_name[:60]}"
    total_found = sum(len(v) for v in section_pdfs.values())
    print(f"  Found {total_found} PDF(s) across {len(section_pdfs)} section(s)")

    downloaded = 0
    skipped = 0

    for section, pdfs in section_pdfs.items():
        print(f"\n  [{section}]")
        for pdf_url, filename in pdfs:
            dest = proc_folder / section / filename
            status = ""

            if dry_run:
                print(f"    [DRY] {filename}  ←  {pdf_url}")
                status = "dry_run"
            elif skip_existing and dest.exists():
                print(f"    [SKIP] {filename} (already exists)")
                skipped += 1
                status = "skipped"
            else:
                print(f"    [↓] {filename} ...", end=" ", flush=True)
                ok = download_pdf(pdf_url, dest, session)
                if ok:
                    size_kb = dest.stat().st_size // 1024
                    print(f"OK ({size_kb} kB)")
                    downloaded += 1
                    status = "downloaded"
                else:
                    status = "error"
                time.sleep(0.3)  # polite delay

            if log_rows is not None:
                log_rows.append({
                    "procedure_id": procedure_id,
                    "drug": drug_name,
                    "section": section,
                    "filename": filename,
                    "url": pdf_url,
                    "local_path": str(dest) if not dry_run else "",
                    "status": status,
                })

    return {
        "id": procedure_id,
        "status": "ok",
        "downloaded": downloaded,
        "skipped": skipped,
        "total_found": total_found,
    }


def write_log(log_rows: list, output_dir: Path):
    """Write a CSV manifest of all processed PDFs."""
    log_path = output_dir / "download_manifest.csv"
    if not log_rows:
        return
    fieldnames = ["procedure_id", "drug", "section", "filename", "url", "local_path", "status"]
    with open(log_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(log_rows)
    print(f"\n  Manifest saved → {log_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Download all PDFs from G-BA Nutzenbewertung procedure pages."
    )
    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--ids", nargs="+", type=int, metavar="ID",
                       help="One or more procedure IDs (e.g. --ids 1102 1292)")
    group.add_argument("--range", nargs=2, type=int, metavar=("START", "END"),
                       help="Inclusive range of IDs (e.g. --range 1100 1110)")
    group.add_argument("--file", type=str, metavar="FILE",
                       help="Text file with one ID per line")

    parser.add_argument("--output", type=str, default="./gba_pdfs",
                        help="Root output directory (default: ./gba_pdfs)")
    parser.add_argument("--dry-run", action="store_true",
                        help="List PDFs without downloading")
    parser.add_argument("--skip-existing", action="store_true", default=True,
                        help="Skip files that already exist (default: True)")
    parser.add_argument("--no-skip-existing", dest="skip_existing", action="store_false",
                        help="Re-download even if file already exists")
    parser.add_argument("--delay", type=float, default=1.0,
                        help="Seconds to wait between procedure pages (default: 1.0)")

    args = parser.parse_args()

    # Resolve IDs
    if args.ids:
        ids = args.ids
    elif args.range:
        ids = list(range(args.range[0], args.range[1] + 1))
    elif args.file:
        with open(args.file) as f:
            ids = [int(line.strip()) for line in f if line.strip().isdigit()]
    else:
        ids = list(range(160, 499))  # ← default range: change these two numbers

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"G-BA PDF Downloader")
    print(f"  Procedures to process: {len(ids)}")
    print(f"  Output directory:      {output_dir.resolve()}")
    print(f"  Dry run:               {args.dry_run}")

    session = requests.Session()
    session.headers.update(HEADERS)

    log_rows = []
    stats_all = []

    for i, proc_id in enumerate(ids):
        stats = process_procedure(
            proc_id,
            output_dir,
            session,
            dry_run=args.dry_run,
            skip_existing=args.skip_existing,
            log_rows=log_rows,
        )
        stats_all.append(stats)
        if i < len(ids) - 1:
            time.sleep(args.delay)

    # Summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    total_dl = sum(s.get("downloaded", 0) for s in stats_all)
    total_skip = sum(s.get("skipped", 0) for s in stats_all)
    total_found = sum(s.get("total_found", 0) for s in stats_all)
    errors = [s["id"] for s in stats_all if s["status"] != "ok"]
    print(f"  Procedures processed: {len(stats_all)}")
    print(f"  PDFs found:           {total_found}")
    print(f"  PDFs downloaded:      {total_dl}")
    print(f"  PDFs skipped:         {total_skip}")
    if errors:
        print(f"  Failed procedures:    {errors}")

    if not args.dry_run:
        write_log(log_rows, output_dir)

    print(f"\nDone. Files saved to: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
