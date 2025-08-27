# -*- coding: utf-8 -*-
"""
Analyse des PDFs (Touch N Pay) -> CSV
- Extraction via pdftotext (Poppler), fallback OCR via Tesseract (CLI)
- Support FR/EN (bilingue) pour CA et Ventes
- Barre de progression colorée (Rich)
- Résumé final "pro": tableau coloré + panel du fichier export
- Historique conservé dans export_analyse_pdf.csv (append au lieu d'écraser)
"""

import os, re, csv, subprocess, tempfile, sys, time
from pathlib import Path
from datetime import datetime
from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn
from rich.console import Console
from rich.theme import Theme
from rich.table import Table
from rich.panel import Panel
from rich.box import HEAVY

# ---------- UI ----------
console = Console(theme=Theme({
    "ok": "bold green","err": "bold red","info": "bold cyan","warn": "bold yellow",
    "hl": "bold white","dim": "dim",
}))

# ========= CONFIG =========
if getattr(sys, 'frozen', False):
    BASE_DIR = Path.home() / "Documents" / "AnalysePDF"
else:
    BASE_DIR = Path(__file__).resolve().parent

ROOT = BASE_DIR / "Mettre les PDF ICI"
OUT_CSV = BASE_DIR / "export_analyse_pdf.csv"
RESSOURCES_DIR = BASE_DIR / "dist_bundle_ressources"
ROOT.mkdir(parents=True, exist_ok=True)
BASE_DIR.mkdir(parents=True, exist_ok=True)

# Binaries
POPPLER_BIN = RESSOURCES_DIR / "poppler-bin"
PDFTOTEXT = str(POPPLER_BIN / "pdftotext.exe")
PDFTOPPM  = str(POPPLER_BIN / "pdftoppm.exe")
TESSERACT_EXE = str(RESSOURCES_DIR / "tesseract" / "tesseract.exe")
TESSDATA_DIR  = str(RESSOURCES_DIR / "tesseract" / "tessdata")
TESS_LANG = "fra+eng"
ENABLE_OCR_FALLBACK = True

# ========= HELPERS =========
def norm_spaces_keep_lines(s: str) -> str:
    s = s.replace("\r", "")
    return "\n".join(re.sub(r"[ \u00A0]+", " ", ln).rstrip() for ln in s.splitlines())

def squash(s: str) -> str: return re.sub(r"\s+", " ", s).strip()
def get_first(regex, text, flags=0):
    m = re.search(regex, text, flags); return m.group(1).strip() if m else ""

def run_pdftotext(pdf_path: str) -> str:
    if not os.path.isfile(PDFTOTEXT): return ""
    try:
        with tempfile.TemporaryDirectory() as td:
            out_txt = Path(td) / "out.txt"
            cmd = [PDFTOTEXT, "-layout", "-nopgbrk", pdf_path, str(out_txt)]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)
            return out_txt.read_text(encoding="utf-8", errors="ignore")
    except Exception: return ""

def run_tesseract_cli_on_pdf(pdf_path: str) -> str:
    if not ENABLE_OCR_FALLBACK: return ""
    if not os.path.isfile(PDFTOPPM) or not os.path.isfile(TESSERACT_EXE): return ""
    full_text = ""
    try:
        with tempfile.TemporaryDirectory() as td:
            out_prefix = Path(td) / "page"
            subprocess.run([PDFTOPPM, "-png", "-r", "450", pdf_path, str(out_prefix)],
                           check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)
            imgs = sorted(Path(td).glob("page*.png"))
            with Progress(TextColumn("  [info]OCR[/info] {task.completed}/{task.total} page(s)"),
                          BarColumn(bar_width=None, complete_style="green", finished_style="bold green", pulse_style="yellow"),
                          TimeElapsedColumn(), console=console, transient=True) as p_ocr:
                task = p_ocr.add_task("OCR pages", total=len(imgs))
                for i, img in enumerate(imgs, 1):
                    txt_out = Path(td) / f"ocr_{i}"
                    cmd_tess = [TESSERACT_EXE, str(img), str(txt_out),
                                "-l", TESS_LANG, "--psm", "6", "--oem", "1",
                                "--tessdata-dir", TESSDATA_DIR]
                    subprocess.run(cmd_tess, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)
                    part = (txt_out.with_suffix(".txt")).read_text(encoding="utf-8", errors="ignore")
                    full_text += part + "\n"; p_ocr.advance(task)
        return full_text
    except Exception: return ""

# ========= PARSING =========
def parse_header(text: str) -> dict:
    lines = [ln for ln in text.splitlines() if ln.strip()]
    joined = "\n".join(lines); id_val = ""
    for ln in lines[:150]:
        if "TOUCH" in ln.upper(): id_val = squash(ln); id_val = re.sub(r"\s*\d{2}/\d{2}/\d{4}.*$", "", id_val).strip(); break
    date_val = get_first(r"\b(\d{2}/\d{2}/\d{4})\b", joined)
    num_rel  = get_first(r"Num[ée]ro\s+de\s+relev[ée]\s*:\s*([0-9]+)", joined, flags=re.IGNORECASE)
    return {"id": id_val, "date": date_val, "Numéro de relevé": num_rel}

def extract_codes_and_key(text: str) -> dict:
    flat = re.sub(r"\s+", " ", text); out = {}
    for idx, val in re.findall(r"Code\s+gratuit\s+(\d+)\s*:\s*([0-9]+(?:\*\*[0-9]/[0-9])?)", flat, flags=re.IGNORECASE):
        out[f"Code gratuit {idx}"] = val
    k1 = get_first(r"\bkey\s+1\s*:\s*([A-Za-z0-9]+)", flat, flags=re.IGNORECASE)
    if k1: out["key 1"] = k1
    return out

_NUM = r"([0-9][0-9\s.,]*€?)"
def _clean_num(s: str) -> str: return re.sub(r"\s+", "", s or "").replace("€", "")

def grab_triple_multiline(text: str, prefix: str, label: str):
    m = re.search(rf"{prefix}\s+{label}", text, flags=re.IGNORECASE)
    if m: window = text[m.end():m.end()+220]; nums = re.findall(_NUM, window)
    else: return "","",""
    return tuple(_clean_num(x) for x in (nums[:3] + ["", "", ""])[:3])

def parse_blocks(text: str, prefix: str, label_map: dict) -> dict:
    out = {}
    for label_src, col_base in label_map.items():
        x,y,z = grab_triple_multiline(text, prefix, label_src)
        out[f"{col_base}_Cumul"], out[f"{col_base}_Interim"], out[f"{col_base}_Interim2"] = x,y,z
    return out

HEADERS = [
    "id","date","Numéro de relevé",
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
    "Code gratuit 1","Code gratuit 2","Code gratuit 3","Code gratuit 4","Code gratuit 5","Code gratuit 6","Code gratuit 7","key 1"
]

# ========= PIPELINE =========
def process_pdf(pdf_path: Path) -> tuple[dict, bool]:
    raw_text = run_pdftotext(str(pdf_path)) or run_tesseract_cli_on_pdf(str(pdf_path))
    text = norm_spaces_keep_lines(raw_text)
    row = {k: "" for k in HEADERS}
    if not text.strip(): row["id"] = pdf_path.stem; return row, False
    row.update(parse_header(text)); row.update(extract_codes_and_key(text))

    # mapping FR + EN
    ca_map = {
        "Total":"CA total","Espèces":"CA Espece","Cash":"CA Espece",
        "Cashless 1":"CA Cashless1","Cashless turnover 1":"CA Cashless1",
        "Cashless 1 Aztek":"CA Cashless1 Aztek","Aztek cashless turnover 1":"CA Cashless1 Aztek",
        "Cashless 2":"CA Cashless2","Cashless turnover 2":"CA Cashless2",
        "Cashless 2 Aztek":"CA Cashless2 Aztek","Aztek cashless turnover 2":"CA Cashless2 Aztek"
    }
    ventes_map = {
        "Total":"Vente Total","Espèces":"Vente Espece","Cash":"Vente Espece",
        "Cashless 1":"Vente Cashless1","Cashless vends 1":"Vente Cashless1",
        "Cashless 1 Aztek":"Vente Cashless1 Aztek","Aztek cashless vends 1":"Vente Cashless1 Aztek",
        "Cashless 2":"Vente Cashless2","Cashless vends 2":"Vente Cashless2",
        "Cashless 2 Aztek":"Vente Cashless2 Aztek","Aztek cashless vends 2":"Vente Cashless2 Aztek"
    }

    row.update(parse_blocks(text,"CA",ca_map))
    row.update(parse_blocks(text,"Ventes",ventes_map))
    row.update(parse_blocks(text,"Turnover",ca_map)) # anglais
    row.update(parse_blocks(text,"Vends",ventes_map)) # anglais
    return row, True

def print_summary(total, ok, errors, failed_files, out_csv):
    last_update = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    table = Table(title="Résumé de l’analyse", box=HEAVY, show_header=False, expand=True)
    table.add_row("📊 PDF analysés", f"[hl]{total}[/hl]")
    table.add_row("✅ PDF analysés OK", f"[ok]{ok}[/ok]")
    table.add_row("❌ Erreurs", f"[err]{errors}[/err]")
    console.print(table)
    if failed_files:
        tf = Table(title="Fichiers en échec", box=HEAVY, show_header=True, expand=True)
        tf.add_column("Nom du fichier", style="err")
        for name in failed_files: tf.add_row(name)
        console.print(tf)
    panel_text = f"[ok][OK][/ok] Export : [hl]{out_csv}[/hl]\n📅 Mis à jour le [hl]{last_update}[/hl]\n[dim]Historique conservé[/dim]"
    console.print(Panel(panel_text, title="Fichier export", border_style="info"))

def main():
    pdfs = sorted([ROOT/f for f in os.listdir(ROOT) if f.lower().endswith(".pdf")])
    if not pdfs: console.print("[warn][INFO][/warn] Aucun PDF trouvé dans le dossier."); return
    results, failed_files = [], []
    with Progress(TextColumn("[bold blue]🔍 Analyse[/bold blue] {task.fields[filename]}"),
                  BarColumn(bar_width=None, complete_style="green", finished_style="bold green", pulse_style="yellow"),
                  "[progress.percentage]{task.percentage:>3.0f}%",TimeElapsedColumn(),TimeRemainingColumn(),
                  console=console, transient=True) as progress:
        task = progress.add_task("Analyse", total=len(pdfs), filename="")
        for pdf in pdfs:
            progress.update(task, filename=pdf.name)
            try: row, ok = process_pdf(pdf); results.append(row); 
            except Exception: r={k:"" for k in HEADERS}; r["id"]=pdf.stem; results.append(r); failed_files.append(pdf.name)
            finally: progress.advance(task)
    file_exists = OUT_CSV.exists()
    with open(OUT_CSV,"a",newline="",encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        if not file_exists: writer.writeheader(); writer.writerows(results)
    total, errors, ok = len(pdfs), len(failed_files), len(pdfs)-len(failed_files)
    print_summary(total,ok,errors,failed_files,OUT_CSV)
    console.print("\n[info]Cette fenêtre se fermera dans 10 minutes...[/info]"); time.sleep(600)

if __name__=="__main__": main()
