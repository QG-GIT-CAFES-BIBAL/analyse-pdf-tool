# -*- coding: utf-8 -*-
"""
Analyse des PDFs (Touch N Pay) -> CSV (v3 - robuste multi-layout)
- Extraction texte multi-stratégies : Poppler (pdftotext) layout/raw, PyPDF2 (fallback), OCR Tesseract en dernier recours
- Parsing adaptatif par lignes (FR/EN) :
    • Détection des libellés (Total/Espèces/Cash/Cashless 1/2 + variantes Aztek)
    • Colonnes "Cumul/Cumulation" + "Interim" + "Interim 2" même si réparties sur plusieurs lignes
- Anti-bruit : normalisation chiffres (espace, point, virgule, €), séparation fiable des tokens
- Score de complétude : succès si ≥ 6 champs numériques CA/Ventes trouvés (configurable)
- Debug : exports .dbg.txt (texte brut) et .dbg.rows.jsonl (lignes tokenisées) par PDF en cas d’échec (pour affiner vite)
"""

import os, re, csv, subprocess, tempfile, sys, time, json
from pathlib import Path
from datetime import datetime

# UI
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

# Binaries (facultatifs suivant OS)
POPPLER_BIN = RESSOURCES_DIR / "poppler-bin"
PDFTOTEXT = str(POPPLER_BIN / "pdftotext.exe") if os.name == "nt" else "pdftotext"
PDFTOPPM  = str(POPPLER_BIN / "pdftoppm.exe")  if os.name == "nt" else "pdftoppm"
TESSERACT_EXE = str(RESSOURCES_DIR / "tesseract" / "tesseract.exe") if os.name == "nt" else "tesseract"
TESSDATA_DIR  = str(RESSOURCES_DIR / "tesseract" / "tessdata")
TESS_LANG = "fra+eng"
ENABLE_OCR_FALLBACK = True  # mettre False si tu veux éviter l’OCR

# Seuil de réussite : au moins N champs CA/Ventes numériques trouvés
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

# ========= EXTRACTION TEXTE MULTI-STRAT =========
def extract_text_any(pdf_path: str) -> str:
    # 1) pdftotext -layout
    txt = run_pdftotext(pdf_path, "layout")
    if strip_ok(txt): return txt
    # 2) pdftotext "raw"
    txt = run_pdftotext(pdf_path, "raw")
    if strip_ok(txt): return txt
    # 3) PyPDF2
    txt = run_pypdf2(pdf_path)
    if strip_ok(txt): return txt
    # 4) OCR
    txt = run_tesseract_cli_on_pdf(pdf_path)
    return txt

def strip_ok(s: str) -> bool:
    return bool(s and s.strip())

# ========= PARSING =========

NUM_TOKEN = r"-?\d+(?:[.,]\d+)?(?:\s*€)?"
TOKEN_RE = re.compile(NUM_TOKEN)

def clean_num_tok(tok: str) -> str:
    if not tok: return ""
    tok = tok.replace("€","").strip()
    # sépare les espaces internes (ex: "145 142" -> "145","142") géré au niveau tokenisation par findall
    return tok.replace(",", ".")

# Synonymes / variantes
LABELS_CA = {
    # label canonique -> liste de variantes attendues
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

COL_HEADERS = [
    # variantes colonnes (FR/EN) ; ordre attendu: Cumul/Cumulation, Interim, Interim 2
    ["Cumul","Cumulation","Cumulated","Cumulative"],
    ["Interim","Interim."],   # on laisse simple
    ["Interim 2","Interim2","Interim2.","Interim. 2","Interim.2"]
]

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

def match_any(token: str, variants: list[str]) -> bool:
    t = token.lower()
    for v in variants:
        if t == v.lower(): return True
        # tolère sans accents / casse partielle
        if v.lower() in t: return True
    return False

def find_col_header_positions(lines_norm: list[str]) -> tuple[int|None, list[int]]:
    """
    Cherche la ligne d'entête contenant Cumul/Cumulation + Interim (+/- Interim 2).
    Retourne (index_ligne, [positions numériques trouvées ensuite] -> ici on ne se sert que de l'index_ligne)
    """
    for i, ln in enumerate(lines_norm[:200]):  # header plutôt haut
        low = ln.lower()
        ok_cumul = any(x.lower() in low for x in COL_HEADERS[0])
        ok_inter = any(x.lower() in low for x in COL_HEADERS[1])
        if ok_cumul and ok_inter:
            return i, []
    return None, []

def parse_numbers_from_near(lines_norm: list[str], start_idx: int, max_span: int = 6) -> list[str]:
    """
    Récupère jusqu'à 3 nombres à partir de start_idx sur quelques lignes suivantes.
    Prend les 3 premiers tokens numériques distincts rencontrés.
    """
    nums = []
    for j in range(start_idx, min(start_idx + max_span, len(lines_norm))):
        for tok in TOKEN_RE.findall(lines_norm[j]):
            val = clean_num_tok(tok)
            if val:
                nums.append(val)
                if len(nums) >= 3:
                    return nums[:3]
    return nums[:3]

def build_reverse_index(label_map: dict) -> dict:
    """
    Construit un dict variante_normalisée -> label canonique
    """
    idx = {}
    for canon, vars in label_map.items():
        for v in vars:
            idx[v.lower()] = canon
    return idx

IDX_CA = build_reverse_index(LABELS_CA)
IDX_V = build_reverse_index(LABELS_VENTES)

def canonical_from_line(ln: str, is_ca: bool) -> str|None:
    low = ln.lower()
    idx = IDX_CA if is_ca else IDX_V
    # essaie exact contient
    for var, canon in idx.items():
        if var in low:
            return canon
    # essais bonus pour “Cashless 1/2 Aztek”
    if "cashless" in low and "1" in low and "aztek" in low:
        return ("CA Cashless1 Aztek" if is_ca else "Vente Cashless1 Aztek")
    if "cashless" in low and "2" in low and "aztek" in low:
        return ("CA Cashless2 Aztek" if is_ca else "Vente Cashless2 Aztek")
    if ("cashless" in low and "1" in low):
        return ("CA Cashless1" if is_ca else "Vente Cashless1")
    if ("cashless" in low and "2" in low):
        return ("CA Cashless2" if is_ca else "Vente Cashless2")
    if any(w in low for w in ["espèces","especes","cash "]):
        return ("CA Espece" if is_ca else "Vente Espece")
    if "total" in low:
        return ("CA total" if is_ca else "Vente Total")
    return None

def parse_blocks_adaptive(text: str) -> dict:
    """
    Lecture par lignes :
    - détecte entête colonnes (Cumul/Cumulation | Interim | Interim 2)
    - pour chaque ligne portant un libellé connu (CA/VENTES), récupère 1 à 3 chiffres dans les lignes proches
    - si pas d'entête colonnes trouvée, on scanne quand même 2-3 lignes autour
    """
    out = {}
    lines = [ln for ln in text.splitlines() if ln.strip()]
    lines_norm = [norm_spaces_keep_lines(ln) for ln in lines]

    # header colonnes (optionnel)
    header_idx, _ = find_col_header_positions(lines_norm)

    def feed_section(is_ca: bool):
        for i, ln in enumerate(lines_norm):
            canon = canonical_from_line(ln, is_ca=is_ca)
            if not canon:
                continue
            # point de départ pour chercher chiffres : si header connu, prends à partir de la même ligne, sinon ligne courante
            start = i if header_idx is None else max(i, header_idx)
            nums = parse_numbers_from_near(lines_norm, start, max_span=6)
            a, b, c = (nums + ["", "", ""])[:3]
            out[f"{canon}_Cumul"] = a
            out[f"{canon}_Interim"] = b
            out[f"{canon}_Interim2"] = c

    # passe CA puis Ventes
    feed_section(is_ca=True)
    feed_section(is_ca=False)
    return out

def process_pdf(pdf_path: Path) -> tuple[dict, bool]:
    # 1) extraction texte
    raw_text = extract_text_any(str(pdf_path))
    text = norm_spaces_keep_lines(raw_text)
    row = {k: "" for k in HEADERS}
    if not text.strip():
        row["id"] = pdf_path.stem
        return row, False

    # 2) entête + codes
    row.update(parse_header(text))
    row.update(extract_codes_and_key(text))

    # 3) parsing adaptatif
    parsed = parse_blocks_adaptive(text)
    row.update(parsed)

    # 4) score de complétude (N champs CA/Vente numériques)
    def _is_num(s):
        try:
            float(s)
            return True
        except:
            return False

    numeric_count = sum(1 for k, v in row.items()
                        if (k.startswith("CA ") or k.startswith("Vente ")) and _is_num(v))
    ok = numeric_count >= MIN_NUMERIC_FIELDS

    # 5) debug si KO
    if not ok:
        dbg_txt = pdf_path.with_suffix(".dbg.txt")
        dbg_rows = pdf_path.with_suffix(".dbg.rows.jsonl")
        try:
            dbg_txt.write_text(text, encoding="utf-8", errors="ignore")
            with dbg_rows.open("w", encoding="utf-8") as f:
                for ln in text.splitlines():
                    tokens = [clean_num_tok(x) for x in TOKEN_RE.findall(ln)]
                    json.dump({"line": ln, "num_tokens": tokens}, f, ensure_ascii=False)
                    f.write("\n")
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
        console.print("[dim]Des fichiers .dbg.txt et .dbg.rows.jsonl ont été générés à côté des PDFs en échec.[/dim]")
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
