# -*- coding: utf-8 -*-
"""
Analyse des PDFs (Touch N Pay) -> CSV
- Extraction via pdftotext (Poppler), fallback OCR via Tesseract (CLI)
- Barre de progression colorée (Rich)
- Résumé final "pro": tableau coloré + panel du fichier export
"""

import os
import re
import csv
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

# ---------- UI (Rich) ----------
from rich.progress import (
    Progress, BarColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn
)
from rich.console import Console
from rich.theme import Theme
from rich.table import Table
from rich.panel import Panel
from rich.box import HEAVY

# Thème lisible
console = Console(theme=Theme({
    "ok": "bold green",
    "err": "bold red",
    "info": "bold cyan",
    "warn": "bold yellow",
    "hl": "bold white",
    "dim": "dim",
}))

# ========= CONFIG DYNAMIQUE =========
# Dossier "AnalysePDF" placé à côté de l'exe (ou du script)
BASE_DIR = Path(__file__).resolve().parent.parent / "AnalysePDF"

ROOT = BASE_DIR / "Mettre les PDF ICI"        # Dossier d'entrée pour les PDF
OUT_CSV = BASE_DIR / "export_analyse_pdf.csv" # Fichier de sortie CSV

# Poppler (pdftotext.exe / pdftoppm.exe)
POPPLER_BIN = Path(__file__).resolve().parent / "dist_bundle_ressources" / "poppler-bin"
PDFTOTEXT   = str(POPPLER_BIN / "pdftotext.exe")
PDFTOPPM    = str(POPPLER_BIN / "pdftoppm.exe")

# Tesseract (CLI)
TESSERACT_EXE = str(Path(__file__).resolve().parent / "dist_bundle_ressources" / "tesseract" / "tesseract.exe")
TESSDATA_DIR  = str(Path(__file__).resolve().parent / "dist_bundle_ressources" / "tesseract" / "tessdata")
TESS_LANG     = "fra+eng"

# Activer ou non l’OCR si pdftotext échoue
ENABLE_OCR_FALLBACK = True


# Poppler (pdftotext.exe / pdftoppm.exe)
POPPLER_BIN = BASE_DIR / "dist_bundle_ressources" / "poppler-bin"
PDFTOTEXT   = str(POPPLER_BIN / "pdftotext.exe")
PDFTOPPM    = str(POPPLER_BIN / "pdftoppm.exe")

# Tesseract (CLI)
TESSERACT_EXE = str(BASE_DIR / "dist_bundle_ressources" / "tesseract" / "tesseract.exe")
TESSDATA_DIR  = str(BASE_DIR / "dist_bundle_ressources" / "tesseract" / "tessdata")
TESS_LANG     = "fra+eng"

# Activer ou non l’OCR si pdftotext échoue
ENABLE_OCR_FALLBACK = True

# ========= HELPERS =========
def norm_spaces_keep_lines(s: str) -> str:
    s = s.replace("\r", "")
    return "\n".join(re.sub(r"[ \u00A0]+", " ", ln).rstrip() for ln in s.splitlines())

def squash(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def get_first(regex, text, flags=0):
    m = re.search(regex, text, flags)
    return m.group(1).strip() if m else ""

def run_pdftotext(pdf_path: str) -> str:
    if not os.path.isfile(PDFTOTEXT):
        return ""
    try:
        with tempfile.TemporaryDirectory() as td:
            out_txt = Path(td) / "out.txt"
            cmd = [PDFTOTEXT, "-layout", "-nopgbrk", pdf_path, str(out_txt)]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)
            return out_txt.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""

def run_tesseract_cli_on_pdf(pdf_path: str) -> str:
    if not ENABLE_OCR_FALLBACK:
        return ""
    if not os.path.isfile(PDFTOPPM) or not os.path.isfile(TESSERACT_EXE):
        return ""
    full_text = ""
    try:
        with tempfile.TemporaryDirectory() as td:
            out_prefix = Path(td) / "page"
            # PDF -> PNG @ 450 dpi
            cmd_ppm = [PDFTOPPM, "-png", "-r", "450", pdf_path, str(out_prefix)]
            subprocess.run(cmd_ppm, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)
            imgs = sorted(Path(td).glob("page*.png"))
            # Mini barre pour les pages OCR
            with Progress(
                TextColumn("  [info]OCR[/info] {task.completed}/{task.total} page(s)"),
                BarColumn(bar_width=None, complete_style="green", finished_style="bold green", pulse_style="yellow"),
                TimeElapsedColumn(),
                console=console,
                transient=True,
            ) as p_ocr:
                task = p_ocr.add_task("OCR pages", total=len(imgs))
                for i, img in enumerate(imgs, 1):
                    txt_out = Path(td) / f"ocr_{i}"
                    cmd_tess = [
                        TESSERACT_EXE, str(img), str(txt_out),
                        "-l", TESS_LANG, "--psm", "6", "--oem", "1",
                        "--tessdata-dir", TESSDATA_DIR
                    ]
                    subprocess.run(cmd_tess, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)
                    part = (txt_out.with_suffix(".txt")).read_text(encoding="utf-8", errors="ignore")
                    full_text += part + "\n"
                    p_ocr.advance(task)
        return full_text
    except Exception:
        return ""

# ========= PARSING EN-TÊTE / CODES =========
def parse_header(text: str) -> dict:
    lines = [ln for ln in text.splitlines() if ln.strip()]
    joined = "\n".join(lines)
    id_val = ""
    for ln in lines[:150]:
        if "TOUCH" in ln.upper():
            id_val = squash(ln)
            id_val = re.sub(r"\s*\d{2}/\d{2}/\d{4}.*$", "", id_val).strip()
            break
    date_val = get_first(r"\b(\d{2}/\d{2}/\d{4})\b", joined)
    num_rel  = get_first(r"Num[ée]ro\s+de\s+relev[ée]\s*:\s*([0-9]+)", joined, flags=re.IGNORECASE)
    return {"id": id_val, "date": date_val, "Numéro de relevé": num_rel}

def extract_codes_and_key(text: str) -> dict:
    flat = re.sub(r"\s+", " ", text)
    out = {}
    for idx, val in re.findall(r"Code\s+gratuit\s+(\d+)\s*:\s*([0-9]+(?:\*\*[0-9]/[0-9])?)", flat, flags=re.IGNORECASE):
        out[f"Code gratuit {idx}"] = val
    k1 = get_first(r"\bkey\s+1\s*:\s*([A-Za-z0-9]+)", flat, flags=re.IGNORECASE)
    if k1:
        out["key 1"] = k1
    return out

# ========= PARSING TABLEAUX =========
_NUM = r"([0-9][0-9\s.,]*€?)"
def _clean_num(s: str) -> str:
    return re.sub(r"\s+", "", s or "").replace("€", "")

def _label_variants(label: str):
    base = label
    base_noacc = (base
        .replace("É","E").replace("é","e")
        .replace("È","E").replace("è","e")
        .replace("Ê","E").replace("ê","e")
        .replace("À","A").replace("à","a")
        .replace("Î","I").replace("î","i")
    )
    return [base, base_noacc] if base_noacc != base else [base]

def _prefix_variants(prefix: str):
    return ["Ventes", "Vente", "VENTES", "VENTE"] if prefix.lower().startswith("vente") else [prefix, prefix.upper()]

def _grab_after(text: str, start_idx: int, max_chars=220):
    window = text[start_idx:start_idx+max_chars]
    mA = re.search(r"Cumul\s+Interim(?:\s*2|\s+2)?\s*\n\s*" + _NUM + r"\s+" + _NUM + r"\s+" + _NUM, window, flags=re.IGNORECASE)
    if mA: return mA.group(1), mA.group(2), mA.group(3)
    mB = re.search(r"Cumul\s+" + _NUM + r".{0,40}?Interim\s+" + _NUM + r".{0,40}?Interim\s*2\s+" + _NUM, window, flags=re.IGNORECASE|re.DOTALL)
    if mB: return mB.group(1), mB.group(2), mB.group(3)
    mC = re.search(_NUM + r"\s+" + _NUM + r"\s+" + _NUM, window, flags=re.IGNORECASE)
    if mC: return mC.group(1), mC.group(2), mC.group(3)
    mD = re.search(r"Cumul\s+" + _NUM + r".{0,40}?Interim\s+" + _NUM, window, flags=re.IGNORECASE|re.DOTALL)
    if mD: return mD.group(1), mD.group(2), ""
    return "", "", ""

def grab_triple_multiline(text: str, prefix: str, label: str):
    for p in _prefix_variants(prefix):
        for lab in _label_variants(label):
            m = re.search(rf"{p}\s+{lab}", text, flags=re.IGNORECASE)
            if m:
                x, y, z = _grab_after(text, m.end())
                if any([x, y, z]):
                    return _clean_num(x), _clean_num(y), _clean_num(z)
    return "","",""

def parse_blocks(text: str, prefix: str, label_map: dict) -> dict:
    out = {}
    for label_src, col_base in label_map.items():
        x,y,z = grab_triple_multiline(text, prefix, label_src)
        out[f"{col_base}_Cumul"]    = x
        out[f"{col_base}_Interim"]  = y
        out[f"{col_base}_Interim2"] = z
    return out

# ========= ENTÊTES CSV =========
HEADERS = [
    "id", "date", "Numéro de relevé",
    "CA total_Cumul","CA total_Interim","CA total_Interim2",
    "CA Espece_Cumul","CA Espece_Interim","CA Espece_Interim2",
    "CA Cashless1_Cumul","CA Cashless1_Interim","CA Cashless1_Interim2",
    "CA Cashless1 Aztek_Cumul","CA Cashless1 Aztek_Interim","CA Cashless1 Aztek_Interim2",
    "CA Cashless2_Cumul","CA Cashless2_Interim","CA Cashless2_Interim2",
    "CA Cashless2 Aztek_Cumul","CA Cashless2 Aztek_Interim","CA Cashless2 Aztek_Interim2",
    "Vente Total_Cumul","Vente Total_Interim","Vente Total_Interim2",
    "Vente Espece_Cumul","Vente Espece_Interim","Vente Espece_Interim2",
    "Vente Cashless1_Cumul","Vente Cashless1_Interim","Vente Cashless1_Interim2",
    "Vente Cashless1 Aztek_Cumul","Vente Cashless1 Aztek_Interim","Vente Cashless1 Aztek_Interim2",
    "Vente Cashless2_Cumul","Vente Cashless2_Interim","Vente Cashless2_Interim2",
    "Vente Cashless2 Aztek_Cumul","Vente Cashless2 Aztek_Interim","Vente Cashless2 Aztek_Interim2",
    "Code gratuit 1","Code gratuit 2","Code gratuit 3","Code gratuit 4","Code gratuit 5","Code gratuit 6","Code gratuit 7",
    "key 1",
]

# ========= PIPELINE =========
def process_pdf(pdf_path: Path) -> tuple[dict, bool]:
    raw_text = run_pdftotext(str(pdf_path))
    if not raw_text.strip():
        raw_text = run_tesseract_cli_on_pdf(str(pdf_path))
    text = norm_spaces_keep_lines(raw_text)

    row = {k: "" for k in HEADERS}
    if not text.strip():
        row["id"] = pdf_path.stem
        return row, False

    row.update(parse_header(text))
    row.update(extract_codes_and_key(text))

    ca_map = {
        "Total": "CA total",
        "Espèces": "CA Espece",
        "Cashless 1": "CA Cashless1",
        "Cashless 1 Aztek": "CA Cashless1 Aztek",
        "Cashless 2": "CA Cashless2",
        "Cashless 2 Aztek": "CA Cashless2 Aztek",
    }
    ventes_map = {
        "Total": "Vente Total",
        "Espèces": "Vente Espece",
        "Cashless 1": "Vente Cashless1",
        "Cashless 1 Aztek": "Vente Cashless1 Aztek",
        "Cashless 2": "Vente Cashless2",
        "Cashless 2 Aztek": "Vente Cashless2 Aztek",
    }

    row.update(parse_blocks(text, "CA", ca_map))
    row.update(parse_blocks(text, "Ventes", ventes_map))

    ok_flag = bool(row.get("id") or row.get("date"))
    if not ok_flag:
        row["id"] = row.get("id") or pdf_path.stem
    return row, ok_flag

def print_summary(total: int, ok: int, errors: int, failed_files: list[str], out_csv: Path):
    last_update = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    # Tableau stats
    table = Table(title="Résumé de l’analyse", box=HEAVY, show_header=False, expand=True)
    table.add_row("📊 PDF analysés", f"[hl]{total}[/hl]")
    table.add_row("✅ PDF analysés OK", f"[ok]{ok}[/ok]")
    table.add_row("❌ Erreurs", f"[err]{errors}[/err]")

    console.print(table)

    # Liste des échecs
    if failed_files:
        tf = Table(title="Fichiers en échec", box=HEAVY, show_header=True, expand=True)
        tf.add_column("Nom du fichier", style="err")
        for name in failed_files:
            tf.add_row(name)
        console.print(tf)

    # Panel export
    panel_text = (
        f"[ok][OK][/ok] Export : [hl]{out_csv}[/hl]\n"
        f"📅 Le fichier [hl]{out_csv.name}[/hl] a été mis à jour le [hl]{last_update}[/hl]\n\n"
        f"[dim]Conseil : Ouvrez le CSV avec Excel / Google Sheets.[/dim]"
    )
    console.print(Panel(panel_text, title="Fichier export", border_style="info"))
    console.print("[magenta]Merci d’avoir utilisé l’outil d’analyse ✅[/magenta]")

def main():
    pdfs = sorted([ROOT / f for f in os.listdir(ROOT) if f.lower().endswith(".pdf")])
    if not pdfs:
        console.print("[warn][INFO][/warn] Aucun PDF trouvé dans le dossier.")
        return

    results = []
    failed_files = []

    with Progress(
        TextColumn("[bold blue]🔍 Analyse[/bold blue] {task.fields[filename]}"),
        BarColumn(bar_width=None, complete_style="green", finished_style="bold green", pulse_style="yellow"),
        "[progress.percentage]{task.percentage:>3.0f}%",
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
        transient=True,
    ) as progress:
        task = progress.add_task("Analyse des PDF", total=len(pdfs), filename="")
        for pdf in pdfs:
            progress.update(task, filename=pdf.name)
            try:
                row, ok = process_pdf(pdf)
                results.append(row)
                if not ok:
                    failed_files.append(pdf.name)
            except Exception:
                r = {k: "" for k in HEADERS}
                r["id"] = Path(pdf).stem
                results.append(r)
                failed_files.append(pdf.name)
            finally:
                progress.advance(task)

    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(results)

    total = len(pdfs)
    errors = len(failed_files)
    ok = total - errors
    print_summary(total, ok, errors, failed_files, OUT_CSV)

if __name__ == "__main__":
    main()
