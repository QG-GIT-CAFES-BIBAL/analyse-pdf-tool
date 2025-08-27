# -*- coding: utf-8 -*-
"""
Analyse des PDFs (Touch N Pay) -> CSV (v5 - double extraction + double parsing + fusion)
- 2 extractions texte complémentaires : pdftotext -layout puis pdftotext -raw (fallback PyPDF2/OCR si besoin)
- 2 parsings complémentaires : fenêtre 400 (rapprochée) puis 800 (élargie)
- Sélection du meilleur résultat par score (nb de champs numériques trouvés) + merge pour combler les trous
- Normalisation chiffres (espace, virgule/point, €)
- Debug auto si score faible
"""

import os, re, csv, subprocess, tempfile, sys, time, json
from pathlib import Path
from datetime import datetime

from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn
from rich.console import Console
from rich.theme import Theme
from rich.table import Table
from rich.panel import Panel
from rich.box import HEAVY

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

# Binaries (selon OS)
POPPLER_BIN = RESSOURCES_DIR / "poppler-bin"
PDFTOTEXT = str(POPPLER_BIN / "pdftotext.exe") if os.name == "nt" else "pdftotext"
PDFTOPPM  = str(POPPLER_BIN / "pdftoppm.exe")  if os.name == "nt" else "pdftoppm"
TESSERACT_EXE = str(RESSOURCES_DIR / "tesseract" / "tesseract.exe") if os.name == "nt" else "tesseract"
TESSDATA_DIR  = str(RESSOURCES_DIR / "tesseract" / "tessdata")
TESS_LANG = "fra+eng"
ENABLE_OCR_FALLBACK = True

# Seuil mini : au moins N champs CA/Ventes numériques pour considérer OK
MIN_NUMERIC_FIELDS = 6

# ========= HELPERS =========
def norm_spaces_keep_lines(s: str) -> str:
    s = s.replace("\r", "")
    return "\n".join(re.sub(r"[ \u00A0]+", " ", ln).rstrip() for ln in s.splitlines())

def squash(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def get_first(regex, text, flags=0):
    m = re.search(regex, text, flags)
    return m.group(1).strip() if m else ""

def _available(cmd: str) -> bool:
    from shutil import which
    return bool(which(cmd)) if os.name != "nt" else os.path.isfile(cmd)

def run_pdftotext(pdf_path: str, mode: str = "layout") -> str:
    if not _available(PDFTOTEXT): return ""
    try:
        with tempfile.TemporaryDirectory() as td:
            out_txt = Path(td) / "out.txt"
            args = ["-layout"] if mode == "layout" else []
            cmd = [PDFTOTEXT, *args, "-nopgbrk", pdf_path, str(out_txt)]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           creationflags=0x08000000 if os.name=="nt" else 0)
            return out_txt.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""

def run_pypdf2(pdf_path: str) -> str:
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(pdf_path)
        pages = []
        for p in reader.pages:
            try:
                pages.append(p.extract_text() or "")
            except Exception:
                pages.append("")
        return "\n".join(pages)
    except Exception:
        return ""

def run_tesseract_cli_on_pdf(pdf_path: str) -> str:
    if not ENABLE_OCR_FALLBACK: return ""
    if not _available(PDFTOPPM) or not _available(TESSERACT_EXE): return ""
    full_text = ""
    try:
        with tempfile.TemporaryDirectory() as td:
            out_prefix = Path(td) / "page"
            subprocess.run([PDFTOPPM, "-png", "-r", "450", pdf_path, str(out_prefix)],
                           check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           creationflags=0x08000000 if os.name=="nt" else 0)
            imgs = sorted(Path(td).glob("page*.png"))
            with Progress(TextColumn("  [info]OCR[/info] {task.completed}/{task.total} page(s)"),
                          BarColumn(bar_width=None, complete_style="green", finished_style="bold green", pulse_style="yellow"),
                          TimeElapsedColumn(), console=console, transient=True) as p_ocr:
                task = p_ocr.add_task("OCR pages", total=len(imgs))
                for i, img in enumerate(imgs, 1):
                    txt_out = Path(td) / f"ocr_{i}"
                    cmd_tess = [TESSERACT_EXE, str(img), str(txt_out),
                                "-l", TESS_LANG, "--psm", "6", "--oem", "1"]
                    if TESSDATA_DIR and os.path.isdir(TESSDATA_DIR):
                        cmd_tess += ["--tessdata-dir", TESSDATA_DIR]
                    subprocess.run(cmd_tess, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   creationflags=0x08000000 if os.name=="nt" else 0)
                    part = (txt_out.with_suffix(".txt")).read_text(encoding="utf-8", errors="ignore")
                    full_text += part + "\n"; p_ocr.advance(task)
        return full_text
    except Exception:
        return ""

def extract_text_strategy(pdf_path: str, strategy: str) -> str:
    """
    strategy: 'layout' | 'raw' | 'pypdf2' | 'ocr'
    """
    if strategy == "layout":
        return run_pdftotext(pdf_path, "layout")
    if strategy == "raw":
        return run_pdftotext(pdf_path, "raw")
    if strategy == "pypdf2":
        return run_pypdf2(pdf_path)
    if strategy == "ocr":
        return run_tesseract_cli_on_pdf(pdf_path)
    return ""

def extract_text_double(pdf_path: str) -> tuple[str, str]:
    """
    Retourne (text1, text2) pour deux passes d'extraction différentes.
    - Pass 1 : pdftotext -layout (fallback pypdf2/ocr si vide)
    - Pass 2 : pdftotext -raw (fallback pypdf2/ocr si vide ET différent de pass1)
    """
    # pass 1
    t1 = extract_text_strategy(pdf_path, "layout")
    if not t1.strip():
        t1 = extract_text_strategy(pdf_path, "pypdf2")
    if not t1.strip():
        t1 = extract_text_strategy(pdf_path, "ocr")

    # pass 2
    t2 = extract_text_strategy(pdf_path, "raw")
    if not t2.strip():
        t2 = extract_text_strategy(pdf_path, "pypdf2")
    if not t2.strip():
        t2 = extract_text_strategy(pdf_path, "ocr")

    # si t2 == t1 (possible), on force une seule passe utile
    if t2.strip() == t1.strip():
        t2 = ""
    return t1, t2

def strip_ok(s: str) -> bool:
    return bool(s and s.strip())

# ========= PARSING =========

# Numéros "propres"
NUM_TOKEN = r"-?\d+(?:[.,]\d+)?(?:\s*€)?"
TOKEN_RE = re.compile(NUM_TOKEN)

def clean_num_tok(tok: str) -> str:
    if not tok: return ""
    return tok.replace("€","").replace(",", ".").strip()

# Canonicals et variantes (FR + EN)
LABELS_CA = {
    "CA total": ["CA Total","Total CA","Total","Total turnover"],
    "CA Espece": ["CA Espèces","CA Espece","Espèces","Cash","Cash turnover"],
    "CA Cashless1": ["Cashless 1","Cashless turnover 1"],
    "CA Cashless1 Aztek": ["Cashless 1 Aztek","Aztek cashless turnover 1","Aztek 1"],
    "CA Cashless2": ["Cashless 2","Cashless turnover 2"],
    "CA Cashless2 Aztek": ["Cashless 2 Aztek","Aztek cashless turnover 2","Aztek 2"],
}
LABELS_VENTES = {
    "Vente Total": ["Ventes Total","Vente Total","Total vends","Total sales"],
    "Vente Espece": ["Ventes Espèces","Vente Espèces","Cash vends","Cash sales","Vends Cash"],
    "Vente Cashless1": ["Ventes Cashless 1","Vente Cashless 1","Cashless vends 1","Cashless sales 1"],
    "Vente Cashless1 Aztek": ["Ventes Cashless 1 Aztek","Vente Cashless 1 Aztek","Aztek cashless vends 1","Aztek sales 1"],
    "Vente Cashless2": ["Ventes Cashless 2","Vente Cashless 2","Cashless vends 2","Cashless sales 2"],
    "Vente Cashless2 Aztek": ["Ventes Cashless 2 Aztek","Vente Cashless 2 Aztek","Aztek cashless vends 2","Aztek sales 2"],
}

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

def parse_header_section(text: str) -> dict:
    lines = [ln for ln in text.splitlines() if ln.strip()]
    joined = "\n".join(lines)
    id_val = ""
    for ln in lines[:150]:
        if "TOUCH" in ln.upper():
            id_val = squash(ln)
            id_val = re.sub(r"\s*\d{2}/\d{2}/\d{4}.*$", "", id_val).strip()
            break
    date_val = get_first(r"\b(\d{2}/\d{2}/\d{4})\b", joined)
    num_rel  = get_first(r"(?:Num[ée]ro\s+de\s+relev[ée]|Report\s+number)\s*:\s*([0-9]+)", joined, flags=re.IGNORECASE)
    return {"id": id_val, "date": date_val, "Numéro de relevé": num_rel}

def extract_codes_and_key(text: str) -> dict:
    flat = re.sub(r"\s+", " ", text)
    out = {}
    for idx, val in re.findall(r"(?:Code\s+gratuit|Free\s+code)\s+(\d+)\s*:\s*([0-9]+(?:\*\*[0-9]/[0-9])?)", flat, flags=re.IGNORECASE):
        out[f"Code gratuit {idx}"] = val
    k1 = get_first(r"\bkey\s*1\s*:\s*([A-Za-z0-9]+)", flat, flags=re.IGNORECASE)
    if k1: out["key 1"] = k1
    return out

def to_flat(text: str) -> str:
    return re.sub(r"[ \u00A0]+", " ", text)

def find_numbers_ahead(flat: str, start_pos: int, max_chars: int = 400, max_tokens: int = 3) -> list[str]:
    window = flat[start_pos : start_pos + max_chars]
    toks = [clean_num_tok(t) for t in TOKEN_RE.findall(window)]
    return [t for t in toks if t][:max_tokens]

def compile_variants_map(label_map: dict) -> list[tuple[str, str, re.Pattern]]:
    res = []
    for canon, variants in label_map.items():
        for v in variants:
            pat = re.compile(rf"(?<!free\s)({re.escape(v)})", re.IGNORECASE)
            res.append((canon, v, pat))
    return res

VARIANTS = compile_variants_map(LABELS_CA) + compile_variants_map(LABELS_VENTES)

def parse_blocks_stream(text: str, win_chars: int = 400) -> dict:
    out = {}
    flat = to_flat(text)
    seen = set()
    for canon, var, pat in VARIANTS:
        if canon in seen: 
            continue
        m = pat.search(flat)
        if not m:
            continue
        nums = find_numbers_ahead(flat, m.end(), max_chars=win_chars, max_tokens=3)
        a, b, c = (nums + ["", "", ""])[:3]
        if any(x for x in (a, b, c)):
            out[f"{canon}_Cumul"] = a
            out[f"{canon}_Interim"] = b
            out[f"{canon}_Interim2"] = c
            seen.add(canon)
    return out

def numeric_score(d: dict) -> int:
    def _is_num(s):
        try: float(s); return True
        except: return False
    return sum(1 for k, v in d.items()
               if (k.startswith("CA ") or k.startswith("Vente ")) and _is_num(v))

def smart_merge(a: dict, b: dict) -> dict:
    """Fusionne deux parsings: on garde la valeur de a, mais si vide dans a et présente dans b -> on prend b"""
    out = dict(a)
    for k, v in b.items():
        if (k not in out) or (not out[k] and v):
            out[k] = v
    return out

def process_pdf(pdf_path: Path) -> tuple[dict, bool]:
    row = {k: "" for k in HEADERS}

    # 1) extractions texte (2 stratégies)
    t1, t2 = extract_text_double(str(pdf_path))
    t1 = norm_spaces_keep_lines(t1)
    t2 = norm_spaces_keep_lines(t2)

    if not strip_ok(t1) and not strip_ok(t2):
        row["id"] = pdf_path.stem
        return row, False

    # 2) entête + codes (depuis la meilleure source dispo)
    src_header = t1 if strip_ok(t1) else t2
    row.update(parse_header_section(src_header))
    row.update(extract_codes_and_key(src_header))

    # 3) parsing par flux (2 fenêtres * 2 textes)
    parsed_variants = []

    if strip_ok(t1):
        p1_t1 = parse_blocks_stream(t1, win_chars=400)
        p2_t1 = parse_blocks_stream(t1, win_chars=800)
        parsed_variants += [p1_t1, p2_t1]

    if strip_ok(t2):
        p1_t2 = parse_blocks_stream(t2, win_chars=400)
        p2_t2 = parse_blocks_stream(t2, win_chars=800)
        parsed_variants += [p1_t2, p2_t2]

    # 4) choisir le meilleur puis fusionner intelligemment pour combler les trous
    if parsed_variants:
        best = max(parsed_variants, key=numeric_score)
        # merge toutes les autres variantes pour compléter
        merged = dict(best)
        for pv in parsed_variants:
            if pv is best: 
                continue
            merged = smart_merge(merged, pv)
        row.update(merged)

    # 5) score de complétude
    ok = numeric_score(row) >= MIN_NUMERIC_FIELDS

    # 6) debug si KO
    if not ok:
        try:
            if strip_ok(t1):
                pdf_path.with_suffix(".dbg.pass1.txt").write_text(t1, encoding="utf-8", errors="ignore")
            if strip_ok(t2):
                pdf_path.with_suffix(".dbg.pass2.txt").write_text(t2, encoding="utf-8", errors="ignore")
            with pdf_path.with_suffix(".dbg.best.json").open("w", encoding="utf-8") as f:
                json.dump({k: v for k, v in row.items() if k.startswith(("CA ", "Vente "))}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    return row, ok

# ========= SORTIE / UI =========
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
        console.print("[dim]Des fichiers .dbg.pass1.txt / .dbg.pass2.txt et .dbg.best.json ont été générés à côté des PDFs en échec.[/dim]")
    panel_text = f"[ok][OK][/ok] Export : [hl]{out_csv}[/hl]\n📅 Mis à jour le [hl]{last_update}[/hl]\n[dim]Historique conservé[/dim]"
    console.print(Panel(panel_text, title="Fichier export", border_style="info"))

def main():
    pdfs = sorted([ROOT/f for f in os.listdir(ROOT) if f.lower().endswith(".pdf")])
    if not pdfs:
        console.print("[warn][INFO][/warn] Aucun PDF trouvé dans le dossier.")
        return

    results, failed_files = [], []
    with Progress(
        TextColumn("[bold blue]🔍 Analyse[/bold blue] {task.fields[filename]}"),
        BarColumn(bar_width=None, complete_style="green", finished_style="bold green", pulse_style="yellow"),
        "[progress.percentage]{task.percentage:>3.0f}%",
        TimeElapsedColumn(),TimeRemainingColumn(),
        console=console, transient=True
    ) as progress:
        task = progress.add_task("Analyse", total=len(pdfs), filename="")
        for pdf in pdfs:
            progress.update(task, filename=pdf.name)
            try:
                row, ok = process_pdf(pdf)
                results.append(row)
                if not ok:
                    failed_files.append(pdf.name)
            except Exception:
                r = {k:"" for k in HEADERS}
                r["id"] = pdf.stem
                results.append(r)
                failed_files.append(pdf.name)
            finally:
                progress.advance(task)

    file_exists = OUT_CSV.exists()
    with open(OUT_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        if not file_exists:
            writer.writeheader()
        writer.writerows(results)

    total, errors, ok = len(pdfs), len(failed_files), len(pdfs)-len(failed_files)
    print_summary(total, ok, errors, failed_files, OUT_CSV)

    console.print("\n[bold green]Merci d'avoir utilisé l'outil d'analyse PDF ![/bold green]")
    console.print("[cyan]Cette fenêtre se fermera automatiquement dans 10 minutes.[/cyan]")
    console.print("[dim]Vous pouvez également la fermer manuellement en cliquant sur la croix.[/dim]\n")
    time.sleep(600)

if __name__=="__main__":
    main()
