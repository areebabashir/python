#!/usr/bin/env python3
"""
PDF / P&ID TAG EXTRACTOR  ->  EXCEL      (multi-file / ZIP version)
==========================================================================
Web app:  python app.py
Run:   python pdf_tag_extractor_multi.py                       (fully interactive)
       python pdf_tag_extractor_multi.py "my.pdf"              (one PDF)
       python pdf_tag_extractor_multi.py "drawings.zip"        (ZIP with many PDFs)
       python pdf_tag_extractor_multi.py "C:\\PIDs"             (folder with many PDFs)
       python pdf_tag_extractor_multi.py a.pdf b.pdf c.zip     (any mix)

What it does
  1. Collects every PDF: a single PDF, all PDFs inside a ZIP (sub-folders too),
     or all PDFs inside a folder (sub-folders too).
  2. Analyses the WHOLE of every PDF (every page, including text inside drawings).
     Instrument bubbles such as "PIT" + "01043" (two separate pieces of text) are
     re-assembled into PIT-01043.  Line numbers, equipment tags and every other
     text line are indexed as well.  Scanned pages use OCR (if Tesseract is installed).
  3. Asks WHICH tags / values to look for (several, separated by commas).
  4. Asks WHICH columns you want in the result (several, separated by commas).
  5. Writes ONE Excel file with the matches from ALL files, including
     "File Name" and "Page" for every match.  You can search again without
     re-analysing.

How to type the tags (comma separated)
  PIT-01043      exact tag (also finds PIT-01043A, PIT-01043B ...)
  PIT            every instrument whose function is PIT  (any number)
  01043          every item containing 01043  (PIT-01043, TIT-01043, TE-01043 ...)
  PIT-010*       wildcard   (* = anything, ? = one character)
  866-E-3303B    equipment tag
  3CA1P01        piping class / any text (searched in every text line)
  re:^\\d+"-RP-  regular expression (start with  re: )

Install once:
  pip install -r requirements.txt
  (optional, only for scanned pages)  pip install pytesseract pillow   + install Tesseract-OCR
"""

import argparse
import fnmatch
import os
import pickle
import re
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime

try:
    import pymupdf  # PyMuPDF >= 1.24
except ImportError:  # older PyMuPDF
    try:
        import fitz as pymupdf
    except ImportError:
        sys.exit("Missing package.  Run:  pip install -r requirements.txt")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing package.  Run:  pip install -r requirements.txt")

CACHE_VERSION = 4


class AnalysisCancelled(Exception):
    """Raised when a caller stops analysis through the progress callback."""


# --------------------------------------------------------------------------
#  Patterns (edit these if your documents use other tag styles)
# --------------------------------------------------------------------------
RE_FUNC = re.compile(r"^[A-Z]{1,5}$")                               # PIT, TI, ZLS ...
RE_NUM = re.compile(r"^(\d{3,6})(?:\s*-?\s*([A-Z]))?$")             # 01043  /  01044 C
RE_INSTR_ONE_LINE = re.compile(r"^([A-Z]{1,5})[\s-]+(\d{3,6})([A-Z])?$")  # "PIT-01043A"
RE_LINE_NO = re.compile(r"""(?<![\w/\-"])(?:\d+-\d+/\d+|\d+/\d+|\d+)"?\s*-\s*[A-Z]{1,4}\s*-\s*\d{3,6}\s*-\s*\d[A-Z0-9]{3,10}(?:\s*-\s*[A-Z0-9]{1,4})?""")
RE_EQUIP = re.compile(r"(?<![\w-])\d{3}-[A-Z]{1,3}-\d{3,5}(?:\s?[A-Z](?![A-Za-z]))?(?![\w-])")
RE_DWG_NO = re.compile(r"\b\d{3}-[A-Z]-[A-Z]{2}-\d{6}-\d{3}\b")

# --------------------------------------------------------------------------
#  Instrument function decoder (ISA-5.1, approximate - Aramco legend overrides)
# --------------------------------------------------------------------------
FIRST = {"A": "Analysis", "B": "Burner", "C": "Conductivity", "D": "Density", "E": "Voltage",
         "F": "Flow", "G": "Gauging", "H": "Hand", "I": "Current", "J": "Power", "K": "Time",
         "L": "Level", "M": "Moisture", "N": "User choice", "O": "User choice", "P": "Pressure",
         "Q": "Quantity", "R": "Radiation", "S": "Speed/Frequency", "T": "Temperature",
         "U": "Multivariable", "V": "Vibration", "W": "Weight/Force", "X": "Unclassified",
         "Y": "Event/State", "Z": "Position"}
SUCC = {"A": "Alarm", "C": "Controller", "E": "Element", "G": "Glass/Gauge", "H": "High",
        "I": "Indicating", "L": "Low", "M": "Middle", "R": "Recording", "S": "Switch",
        "T": "Transmitter", "V": "Valve", "Y": "Relay/Converter", "Z": "Driver/Actuator",
        "Q": "Integrate/Totalize", "X": "Accessory", "B": "Burner/Flame", "D": "Differential",
        "P": "Test point", "W": "Well", "K": "Control station"}
OVERRIDE = {"ZC": "Shutdown logic", "ZV": "Safety shutdown valve", "ZBV": "Emergency blowdown valve",
            "AOV": "On-off valve", "AIT": "Analyzer transmitter", "XL": "Pilot light",
            "LS": "Limit switch", "ESD": "Emergency shutdown system", "HS": "Hand switch",
            "HV": "Hand valve", "FV": "Flow control valve", "PV": "Pressure control valve",
            "LV": "Level control valve", "TV": "Temperature control valve"}


def decode_function(func):
    if not func:
        return ""
    if func in OVERRIDE:
        return OVERRIDE[func]
    first = FIRST.get(func[0], "?")
    rest = func[1:]
    parts = [first]
    if rest[:1] in ("D", "Q", "F") and len(rest) > 1:
        parts.append({"D": "Differential", "Q": "Quantity", "F": "Ratio"}[rest[0]])
        rest = rest[1:]
    for ch in rest:
        parts.append(SUCC.get(ch, ch))
    return " ".join(parts)


# --------------------------------------------------------------------------
#  Collecting input files (PDF / ZIP / folder)
# --------------------------------------------------------------------------
def extract_zip(zip_path):
    """Extract a ZIP next to itself (re-used on later runs). Returns the folder."""
    dest = os.path.splitext(zip_path)[0] + "_unzipped"
    try:
        os.makedirs(dest, exist_ok=True)
        test = os.path.join(dest, ".w")
        open(test, "w").close()
        os.remove(test)
    except Exception:                                  # read-only location
        dest = tempfile.mkdtemp(prefix="tagzip_")
    base = os.path.realpath(dest)
    with zipfile.ZipFile(zip_path) as z:
        for m in z.infolist():
            if m.is_dir() or m.filename.startswith("__MACOSX") or os.path.basename(m.filename).startswith("._"):
                continue
            target = os.path.realpath(os.path.join(dest, m.filename))
            if not target.startswith(base + os.sep):    # protection against '../' paths
                continue
            if os.path.exists(target) and os.path.getsize(target) == m.file_size:
                continue                                 # already extracted
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with z.open(m) as src, open(target, "wb") as out:
                out.write(src.read())
    return dest


def pdfs_in_folder(folder):
    found = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".pdf") and not f.startswith("._"):
                found.append(os.path.join(root, f))
    return sorted(found, key=lambda p: p.lower())


def collect_pdfs(inputs):
    """inputs: list of paths (pdf / zip / folder). Returns [(display_name, abs_path), ...]."""
    result, seen = [], set()

    def add(display, path):
        key = os.path.realpath(path)
        if key not in seen:
            seen.add(key)
            result.append((display, os.path.abspath(path)))

    for p in inputs:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            for f in pdfs_in_folder(p):
                add(os.path.relpath(f, p), f)
        elif p.lower().endswith(".zip") and os.path.isfile(p):
            print(f"  Unzipping {os.path.basename(p)} ...")
            folder = extract_zip(p)
            for f in pdfs_in_folder(folder):
                add(os.path.relpath(f, folder), f)
        elif p.lower().endswith(".pdf") and os.path.isfile(p):
            add(os.path.basename(p), p)
        else:
            print(f"  [!] Skipped (not a PDF / ZIP / folder): {p}")

    # make display names unique (same file name in different sub-folders is already
    # distinguished by its relative path; this handles identical names from separate inputs)
    names = {}
    for i, (d, p) in enumerate(result):
        names.setdefault(d, []).append(i)
    for d, idxs in names.items():
        if len(idxs) > 1:
            for i in idxs:
                result[i] = (f"{os.path.basename(os.path.dirname(result[i][1]))}/{d}", result[i][1])
    return result


# --------------------------------------------------------------------------
#  Analysis
# --------------------------------------------------------------------------
def visual_rect(page, bbox):
    """bbox in PDF space -> rect as the page is displayed (handles rotated pages)."""
    return pymupdf.Rect(bbox) * page.rotation_matrix


def read_title_block(page, spans):
    """Drawing number / title / revision from the bottom-right title block."""
    W, H = page.rect.width, page.rect.height
    blk = [s for s in spans if s["r"].x0 > W * 0.80 and s["r"].y0 > H * 0.85]
    info = {"dwg": "", "title": "", "rev": ""}
    hdr = next((s for s in blk if s["t"] == "DRAWING NUMBER"), None)
    if hdr:
        row = sorted([s for s in blk if 8 < s["r"].y0 - hdr["r"].y0 < 30], key=lambda s: s["r"].x0)
        vals = [s["t"] for s in row]
        if len(vals) >= 5:
            info["dwg"] = "-".join(vals[:4])
            info["rev"] = vals[4]
        elif vals:
            info["dwg"] = "-".join(vals)
    pid = next((s for s in blk if "INSTRUMENTATION DIAGRAM" in s["t"] or "DIAGRAM" in s["t"]), None)
    if pid:
        below = sorted([s for s in blk if 10 < s["r"].y0 - pid["r"].y0 < 40 and s["t"] != "PID"],
                       key=lambda s: s["r"].y0)
        if below:
            info["title"] = max(below, key=lambda s: len(s["t"]))["t"]
    if not info["dwg"]:
        for s in blk:
            m = RE_DWG_NO.search(s["t"])
            if m:
                info["dwg"] = m.group(0)
                break
    return info


def page_spans(page):
    spans, lines = [], []
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        for l in b.get("lines", []):
            parts = []
            for s in l.get("spans", []):
                t = s["text"].strip()
                if not t:
                    continue
                sp = {"t": t, "r": visual_rect(page, s["bbox"]), "size": s["size"]}
                spans.append(sp)
                parts.append(sp)
                sp["i"] = len(spans) - 1
            if parts:
                txt = " ".join(p["t"] for p in parts)
                r = pymupdf.Rect(parts[0]["r"])
                for p in parts[1:]:
                    r |= p["r"]
                lines.append({"t": txt, "r": r, "idx": [p["i"] for p in parts]})
    return spans, lines


def pair_instruments(spans):
    """Join 'PIT' (above) + '01043' (below) -> instrument tags. Returns (tags, used_span_ids)."""
    funcs = [(i, s) for i, s in enumerate(spans) if RE_FUNC.match(s["t"])]
    nums = [(i, s) for i, s in enumerate(spans) if RE_NUM.match(s["t"])]
    cand = []
    for fi, f in funcs:
        fr = f["r"]
        h = max(fr.height, 5)
        fcx = (fr.x0 + fr.x1) / 2
        for ni, n in nums:
            nr = n["r"]
            dy = nr.y0 - fr.y1
            if not (-0.4 * h <= dy <= 0.9 * h):
                continue
            dx = abs((nr.x0 + nr.x1) / 2 - fcx)
            if dx > 0.9 * h + 4:
                continue
            cand.append((dx + abs(dy), fi, ni))
    cand.sort()
    used_f, used_n, out = set(), set(), []
    for _, fi, ni in cand:
        if fi in used_f or ni in used_n:
            continue
        used_f.add(fi)
        used_n.add(ni)
        f, n = spans[fi], spans[ni]
        m = RE_NUM.match(n["t"])
        out.append((f["t"], m.group(1), m.group(2) or "", f["r"] | n["r"]))
    return out, used_f | used_n


_WIN_OCR_PS = r"""
param([string]$InPath, [string]$OutPath)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq "IAsyncOperation``1"
})[0]
function Await($WinRtTask, $ResultType) {
  $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  $netTask.Result
}
$null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime]
$lang = New-Object Windows.Globalization.Language "en-US"
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if (-not $engine) { throw "Windows OCR has no English language" }
function Read-Words([string]$Path) {
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $converted = [Windows.Graphics.Imaging.SoftwareBitmap]::Convert($bitmap, [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8, [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)
  $result = Await ($engine.RecognizeAsync($converted)) ([Windows.Media.Ocr.OcrResult])
  $rows = New-Object System.Collections.Generic.List[object]
  foreach ($line in $result.Lines) {
    foreach ($w in $line.Words) {
      $r = $w.BoundingRect
      $rows.Add([pscustomobject]@{ T = $w.Text; X = $r.X; Y = $r.Y; W = $r.Width; H = $r.Height })
    }
  }
  return $rows
}
function Score($rows) {
  $n = 0
  foreach ($w in $rows) { if (($w.T -replace "[^A-Za-z0-9]", "").Length -ge 4) { $n++ } }
  return $n
}
$src = New-Object System.Drawing.Bitmap $InPath
$flips = @{
  "0" = $null
  "90" = [System.Drawing.RotateFlipType]::Rotate90FlipNone
  "180" = [System.Drawing.RotateFlipType]::Rotate180FlipNone
  "270" = [System.Drawing.RotateFlipType]::Rotate270FlipNone
}
$bestScore = -1
$bestRot = "0"
$dir = Split-Path $OutPath
foreach ($name in @("0","90","180","270")) {
  $trial = New-Object System.Drawing.Bitmap $src
  if ($flips[$name] -ne $null) { $trial.RotateFlip($flips[$name]) }
  $small = $trial
  $maxEdge = [Math]::Max($trial.Width, $trial.Height)
  if ($maxEdge -gt 1200) {
    $scale = 1200.0 / $maxEdge
    $small = New-Object System.Drawing.Bitmap $trial, ([int]($trial.Width * $scale)), ([int]($trial.Height * $scale))
  }
  $tmp = Join-Path $dir ("ocr-" + $name + ".png")
  $small.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png)
  if (-not [object]::ReferenceEquals($small, $trial)) { $small.Dispose() }
  $trial.Dispose()
  $score = Score (Read-Words $tmp)
  if ($score -gt $bestScore) { $bestScore = $score; $bestRot = $name }
}
$full = New-Object System.Drawing.Bitmap $src
if ($flips[$bestRot] -ne $null) { $full.RotateFlip($flips[$bestRot]) }
$fullPath = Join-Path $dir "ocr-best.png"
$full.Save($fullPath, [System.Drawing.Imaging.ImageFormat]::Png)
$full.Dispose()
$src.Dispose()
$words = Read-Words $fullPath
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("ROT`t$bestRot`t$bestScore")
foreach ($w in $words) {
  $t = ($w.T -replace "[\r\n\t]", " ")
  $lines.Add(("{0}`t{1:0.##}`t{2:0.##}`t{3:0.##}`t{4:0.##}" -f $t, $w.X, $w.Y, $w.W, $w.H))
}
[System.IO.File]::WriteAllLines($OutPath, $lines)
"""


def _merge_ocr_words(words):
    """Join pieces of one cell, such as a tag split at a hyphen."""
    pending = sorted(words, key=lambda w: (w[2], w[1]))
    lines, current, last_y = [], [], None
    for word in pending:
        y = word[2] + word[4] / 2
        if last_y is None or abs(y - last_y) <= max(word[4], 8) * 0.7:
            current.append(word)
            last_y = y if last_y is None else (last_y + y) / 2
        else:
            if current:
                lines.append(current)
            current, last_y = [word], y
    if current:
        lines.append(current)
    merged = []
    for line in lines:
        line.sort(key=lambda w: w[1])
        buf = None
        for text, x, y, w, h in line:
            if buf is None:
                buf = [text, x, y, w, h]
                continue
            gap = x - (buf[1] + buf[3])
            hyphen = buf[0].endswith("-") or text.startswith("-")
            separate_words = buf[0].isalpha() and text.isalpha() and len(buf[0]) >= 2 and len(text) >= 2
            if gap <= max(12, buf[4] * 0.8) and (hyphen or not separate_words):
                buf[0] = buf[0] + text
                right = max(buf[1] + buf[3], x + w)
                bottom = max(buf[2] + buf[4], y + h)
                buf[1] = min(buf[1], x)
                buf[2] = min(buf[2], y)
                buf[3] = right - buf[1]
                buf[4] = bottom - buf[2]
            else:
                merged.append(tuple(buf))
                buf = [text, x, y, w, h]
        if buf:
            merged.append(tuple(buf))
    return merged


def _tidy_ocr_text(text):
    text = re.sub(r"\s+", "", text or "")
    found = re.search(r"(?:\d{3}-){1,2}[A-Z]{1,5}-\d{3,6}[A-Z]{0,3}", text, re.I)
    if found and 0 < len(text) - len(found.group(0)) <= 2:
        return found.group(0)
    return text.strip()


def _normalize_model(value):
    """Windows OCR reads a printed 1 as the letter I inside long item codes."""
    parts = []
    for piece in (value or "").split(" | "):
        if len(piece) >= 12 and piece.isalnum() and any(ch.isdigit() for ch in piece) and "I" in piece:
            piece = piece.replace("I", "1")
        if piece:
            parts.append(piece)
    return " | ".join(parts)


def _ocr_windows(page, dpi=230):
    """Read a scanned page with the Windows OCR engine. Returns (text, rect) or None."""
    pix = page.get_pixmap(dpi=dpi)
    folder = tempfile.mkdtemp(prefix="tagocr-")
    image_path = os.path.join(folder, "page.png")
    words_path = os.path.join(folder, "words.txt")
    script_path = os.path.join(folder, "ocr.ps1")
    pix.save(image_path)
    with open(script_path, "w", encoding="utf-8") as handle:
        handle.write(_WIN_OCR_PS)
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0
    try:
        done = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path,
             "-InPath", image_path, "-OutPath", words_path],
            capture_output=True, text=True, timeout=180, creationflags=flags,
        )
        if done.returncode != 0 or not os.path.exists(words_path):
            return None
        words = []
        with open(words_path, encoding="utf-8") as handle:
            for line in handle:
                parts = line.rstrip("\n").split("\t")
                if len(parts) != 5 or parts[0] == "ROT":
                    continue
                text, x, y, w, h = parts
                if text.strip():
                    words.append((text.strip(), float(x), float(y), float(w), float(h)))
    except Exception:
        return None
    finally:
        for name in os.listdir(folder):
            try:
                os.remove(os.path.join(folder, name))
            except OSError:
                pass
        try:
            os.rmdir(folder)
        except OSError:
            pass
    if not words:
        return None
    k = 72.0 / dpi
    out = []
    for text, x, y, w, h in _merge_ocr_words(words):
        text = _tidy_ocr_text(text)
        if text:
            out.append((text, pymupdf.Rect(x * k, y * k, (x + w) * k, (y + h) * k)))
    return out


def ocr_lines(page, dpi=230):
    """OCR fall-back for scanned / raster content. Returns list of (text, rect)."""
    try:
        import pytesseract
        from PIL import Image
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, config="--psm 11")
        k = 72.0 / dpi
        out = []
        for i, t in enumerate(data["text"]):
            t = t.strip()
            if t and float(data["conf"][i]) > 40:
                x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
                out.append((t, pymupdf.Rect(x * k, y * k, (x + w) * k, (y + h) * k)))
        if out:
            return out
    except Exception:
        pass
    if sys.platform == "win32":
        found = _ocr_windows(page, dpi)
        if found:
            return found
    return None


def analyze_pdf(path, ocr_mode="auto", label="", progress=None):
    """Returns dict(pages=[meta...], items=[tuple...])
    item = (kind, text, func, loop_no, suffix, page_no, x, y, source, full_line)

    progress, when set, is called once per page with a dict. Raise AnalysisCancelled
    from that callback to stop. The PDF is closed either way.
    """
    doc = pymupdf.open(path)
    n = doc.page_count
    pages, items = [], []
    t0 = time.time()
    ocr_missing_warned = False
    short = (label if len(label) <= 38 else "..." + label[-35:])
    try:
        for pno in range(n):
            page = doc[pno]
            spans, lines = page_spans(page)
            meta = read_title_block(page, spans)
            meta["page"] = pno + 1
            pages.append(meta)

            # ---- decide on OCR -------------------------------------------------
            use_ocr = ocr_mode == "all"
            if ocr_mode == "auto":
                img_area = 0.0
                try:
                    for img in page.get_images(full=True):
                        for r in page.get_image_rects(img[0]):
                            img_area += r.width * r.height
                except Exception:
                    pass
                use_ocr = (len(spans) < 15) or (img_area > 0.25 * page.rect.width * page.rect.height)

            # ---- instruments (re-assembled bubbles) ----------------------------
            tags, used = pair_instruments(spans)
            for func, num, suf, r in tags:
                items.append(("Instrument", f"{func}-{num}{suf}", func, num, suf, pno + 1,
                              round((r.x0 + r.x1) / 2, 1), round((r.y0 + r.y1) / 2, 1), "text", ""))

            # ---- every text line ------------------------------------------------
            for ln in lines:
                if all(i in used for i in ln["idx"]):
                    continue      # already part of a re-assembled instrument tag
                t = ln["t"]
                r = ln["r"]
                cx, cy = round((r.x0 + r.x1) / 2, 1), round((r.y0 + r.y1) / 2, 1)
                m1 = RE_INSTR_ONE_LINE.match(t)
                if m1:
                    items.append(("Instrument", f"{m1.group(1)}-{m1.group(2)}{m1.group(3) or ''}", m1.group(1),
                                  m1.group(2), m1.group(3) or "", pno + 1, cx, cy, "text", t))
                    continue
                for m in RE_LINE_NO.finditer(t):
                    items.append(("Line No.", re.sub(r"\s*-\s*", "-", m.group(0)).strip(), "", "", "", pno + 1, cx, cy, "text", t))
                for m in RE_EQUIP.finditer(t):
                    items.append(("Equipment", re.sub(r"\s", "", m.group(0)), "", "", "", pno + 1, cx, cy, "text", t))
                items.append(("Text", t, "", "", "", pno + 1, cx, cy, "text", t))

            # ---- OCR for image-only content -------------------------------------
            if use_ocr:
                res = ocr_lines(page)
                if res is None:
                    if not ocr_missing_warned:
                        print("\n  [!] %s page %d looks scanned but pytesseract/Pillow/Tesseract is not installed "
                              "- OCR skipped." % (short, pno + 1))
                        ocr_missing_warned = True
                else:
                    for t, r in res:
                        cx, cy = round((r.x0 + r.x1) / 2, 1), round((r.y0 + r.y1) / 2, 1)
                        kind = "Equipment" if RE_EQUIP.fullmatch(t) else "Line No." if RE_LINE_NO.fullmatch(t) else "Text"
                        items.append((kind, t, "", "", "", pno + 1, cx, cy, "OCR", t))

            el = time.time() - t0
            eta = el / (pno + 1) * (n - pno - 1)
            print(f"\r  {short}  page {pno + 1}/{n}   items: {len(items):,}   ETA {eta:4.0f}s   ", end="", flush=True)
            if progress:
                progress({
                    "page": pno + 1,
                    "pages": n,
                    "items": len(items),
                    "eta": round(eta),
                    "label": short,
                    "cached": False,
                })
        return {"version": CACHE_VERSION, "pages": pages, "items": items}
    finally:
        print()
        doc.close()


def load_or_analyze(path, ocr_mode, use_cache=True, label="", progress=None):
    st = os.stat(path)
    key = (CACHE_VERSION, st.st_size, int(st.st_mtime), ocr_mode)
    cache_file = path + ".tagcache.pkl"
    if use_cache and os.path.exists(cache_file):
        try:
            with open(cache_file, "rb") as f:
                blob = pickle.load(f)
            if blob.get("key") == key:
                print(f"  {label or os.path.basename(path)}: loaded from cache")
                data = blob["data"]
                if progress:
                    progress({
                        "page": len(data["pages"]),
                        "pages": len(data["pages"]),
                        "items": len(data["items"]),
                        "eta": 0,
                        "label": label or os.path.basename(path),
                        "cached": True,
                    })
                return data
        except AnalysisCancelled:
            raise
        except Exception:
            pass
    data = analyze_pdf(path, ocr_mode, label or os.path.basename(path), progress=progress)
    try:
        with open(cache_file, "wb") as f:
            pickle.dump({"key": key, "data": data}, f)
    except Exception:
        pass
    return data


def analyze_all(files, ocr_mode, use_cache=True, progress=None):
    """files: [(display, path)] -> datasets [(display, path, data)]"""
    datasets = []
    nfiles = len(files)
    for i, (display, path) in enumerate(files, 1):
        print(f"\n[{i}/{nfiles}] {display}")

        def cb(info, i=i, display=display):
            if progress:
                merged = dict(info)
                merged["file_index"] = i
                merged["file_count"] = nfiles
                merged["file_name"] = display
                progress(merged)

        try:
            if progress:
                progress({
                    "file_index": i,
                    "file_count": nfiles,
                    "file_name": display,
                    "page": 0,
                    "pages": 0,
                    "items": 0,
                    "eta": 0,
                    "label": display,
                    "cached": False,
                })
            data = load_or_analyze(path, ocr_mode, use_cache, display, progress=cb)
            datasets.append((display, path, data))
        except AnalysisCancelled:
            raise
        except Exception as e:
            print(f"  [!] Could not read {display}: {e}")
    return datasets


# --------------------------------------------------------------------------
#  Searching
# --------------------------------------------------------------------------
def norm(s):
    return re.sub(r"[\s\-_\"'`]", "", s).upper()


def build_matcher(q):
    q = q.strip()
    if q.lower().startswith("re:"):
        rx = re.compile(q[3:], re.I)
        return lambda it: bool(rx.search(it[1]))
    nq = norm(q)
    if "*" in nq or "?" in nq:
        return lambda it: fnmatch.fnmatchcase(norm(it[1]), nq)
    if nq.isalpha() and len(nq) <= 6:             # a function code such as PIT, HS, ZLS
        return lambda it: (it[2] == nq) or (norm(it[1]) == nq)
    if len(nq) < 4:                               # very short -> exact only
        return lambda it: norm(it[1]) == nq
    return lambda it: nq in norm(it[1])


COLUMNS = ["Query", "Tag Number", "Model or Item Code", "Serial No", "Part Number",
           "Reference File", "Related Values", "Tag / Value", "Type", "File Name", "Page",
           "Drawing Number", "Drawing Title", "Revision", "Function", "Function Description",
           "Loop Number", "Suffix", "X (pt)", "Y (pt)", "Nearby Text", "Full Line Text",
           "Source", "File Path"]
DEFAULT_COLUMNS = ["Query", "Tag Number", "Model or Item Code", "Serial No", "Part Number",
                   "Reference File", "File Name", "Page"]
REQUIRED_COLUMNS = ["File Name", "Page"]

RE_DWG_CELL = re.compile(r"^(?:\d{3}-[A-Z]-)?[A-Z]{2}-\d{5,6}-\d{3}$", re.I)
RE_ROW_TAG = re.compile(
    r"^(?:\d{3}-){1,2}[A-Z]{1,5}-\d{3,6}[A-Z]{0,3}$"
    r"|^\d{3}-[A-Z]-\d{3,5}[A-Z]?$"
    r"|^[A-Z]{1,5}-\d{3,6}[A-Z]?$",
    re.I)
RE_ITEM_CODE = re.compile(r"\d[A-Z]{1,4}\d[A-Z]\d{2}", re.I)
_HEADER_WORD = {"SR", "SERIAL", "TAG", "P&ID", "PART", "MODEL", "ITEM", "DRAWING", "DWG", "SL", "SI"}
_HEADER_TAIL = {"NO", "NUMBER", "CODE"}


def _field_of_label(label):
    """Map a table heading to tag, model, serial, part, or the reference drawing."""
    s = re.sub(r"\s+", " ", label).strip().upper().replace("0P&ID", "P&ID")
    s = s.replace(".", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if "SYMBOL" in s or len(s) > 42:
        return ""
    if re.search(r"\bPART\s*(NO|NUMBER)\b|\bP/N\b", s):
        return "part"
    if re.search(r"\bMODEL\b|\bITEM\s*CODE\b", s):
        return "model"
    if re.search(r"\bITEM\s*(NO|NUMBER)\b", s):
        return ""
    if "VALVE" in s and re.search(r"SL|SI|SERIAL|NOS", s):
        return "serial"
    if re.search(r"\bSR\s*NO\b|\bSL\s*N|\bSI\s*N|\bSERIAL\b|\bS/N\b", s):
        return "serial"
    if re.search(r"\bP&ID\s*NO\b|\bDRAWING\s*(NO|NUMBER)\b|\bDWG\s*NO\b", s) or s == "P&ID":
        return "ref"
    if re.search(r"\bTAG\b", s):
        return "tag"
    return ""


def _append_unique(bucket, value):
    value = (value or "").strip()
    if not value or value in ("-", "--", "—"):
        return
    if value not in bucket:
        bucket.append(value)


def _page_text_items(items):
    seen, out = set(), []
    for it in items:
        if it[0] != "Text":
            continue
        key = (it[5], round(it[6], 1), round(it[7], 1), it[1])
        if key in seen:
            continue
        seen.add(key)
        out.append(it)
    return out


def _header_cells(text_items):
    """Column headings, including 'SR.' stacked above 'NO.'."""
    headers, used = [], set()
    for i, it in enumerate(text_items):
        field = _field_of_label(it[1])
        if not field:
            continue
        headers.append({"x": it[6], "y": it[7], "field": field, "label": it[1].strip()})
        used.add(i)
    for i, it in enumerate(text_items):
        if i in used:
            continue
        word = it[1].strip().upper().rstrip(".")
        if word not in _HEADER_WORD:
            continue
        for j, other in enumerate(text_items):
            if j in used or other[5] != it[5]:
                continue
            tail = other[1].strip().upper().rstrip(".")
            if tail not in _HEADER_TAIL:
                continue
            if abs(other[6] - it[6]) < 22 and 0 <= (other[7] - it[7]) <= 24:
                label = it[1].strip() + " " + other[1].strip()
                field = _field_of_label(label)
                if field:
                    headers.append({"x": it[6], "y": other[7], "field": field, "label": label})
                    used.add(i)
                    used.add(j)
                break
    # Headings written across one line, such as "Item" beside "Code".
    for i, it in enumerate(text_items):
        if i in used:
            continue
        word = it[1].strip().upper().rstrip(".")
        if word not in _HEADER_WORD:
            continue
        nearest = None
        for j, other in enumerate(text_items):
            if j == i or j in used or other[5] != it[5]:
                continue
            if abs(other[7] - it[7]) > 10:
                continue
            delta = other[6] - it[6]
            if 8 <= delta <= 90 and (nearest is None or delta < nearest[0]):
                nearest = (delta, j, other)
        if nearest is None:
            continue
        label = it[1].strip() + " " + nearest[2][1].strip()
        field = _field_of_label(label)
        if not field:
            continue
        headers.append({"x": it[6], "y": max(it[7], nearest[2][7]), "field": field, "label": label})
        used.add(i)
        used.add(nearest[1])
    return headers, used


def _tables_from_headers(headers):
    """Side-by-side copies of the same list become separate tables."""
    bands = []
    for header in sorted(headers, key=lambda h: (h["y"], h["x"])):
        if bands and abs(header["y"] - bands[-1][-1]["y"]) <= 36:
            bands[-1].append(header)
        else:
            bands.append([header])
    tables = []
    for band in bands:
        band.sort(key=lambda h: h["x"])
        current, seen = [], set()
        for index, header in enumerate(band):
            field = header["field"]
            later_tag = any(other["field"] == "tag" for other in band[index + 1:])
            # A repeated tag, or a repeated column that begins the next side-by-side list.
            # A second SI / serial column on the same row stays with that row.
            if field != "extra" and field in seen and (field == "tag" or later_tag):
                if len(seen) >= 2 and len(current) >= 2:
                    tables.append(current)
                current, seen = [header], {field}
            else:
                current.append(header)
                if field != "extra":
                    seen.add(field)
        if len(seen) >= 2 and len(current) >= 2:
            tables.append(current)
    return tables


def _item_code_text(text):
    return len(text) >= 12 and text.isalnum() and not re.search(r"\d{2}[A-Z]\d{4,}", text, re.I)


def _serial_value(text):
    """A SI / serial cell, not an item code or a sentence."""
    text = (text or "").strip()
    if (not text) or (" " in text) or len(text) > 24 or _item_code_text(text) or RE_ROW_TAG.match(text):
        return False
    digits = sum(ch.isdigit() for ch in text)
    return digits >= 1 and digits + 3 >= len(re.sub(r"[^A-Za-z0-9]", "", text))


def _pick_serial(pairs):
    """Prefer the SI / valve serial column over an item number or item code."""
    usable = [(label, value) for label, value in pairs if _serial_value(value)]
    if not usable:
        return ""

    def rank(item):
        label, value = item
        heading = re.sub(r"\s+", " ", label).upper()
        valve = 1 if re.search(r"VALVE|OFFERED", heading) else 0
        si = 1 if re.search(r"\b(SL|SI|SR|SERIAL)\b", heading) and "ITEM" not in heading else 0
        patterned = 1 if re.search(r"\d{2}[A-Z]\d{4,}", value, re.I) else 0
        return (patterned, valve, si, len(value))

    return max(usable, key=rank)[1]


def _cell_role(field, text):
    if re.search(r"\d{2}[A-Z]\d{4,}", text or "", re.I) and not RE_ROW_TAG.match(text or ""):
        return "serial"
    if field == "serial" and not _serial_value(text):
        return "related"
    if field in ("tag", "model", "serial", "part", "ref"):
        return field
    if RE_DWG_CELL.match(text):
        return "ref"
    if RE_LINE_NO.search(text) or (RE_ITEM_CODE.search(text) and not RE_ROW_TAG.match(text)):
        return "model"
    if RE_ROW_TAG.match(text):
        return "tag"
    return "related"


def _records_from_table(table, text_items):
    baseline = max(h["y"] for h in table)
    x_min = min(h["x"] for h in table) - 40
    x_max = max(h["x"] for h in table) + 90
    xs = [h["x"] for h in table]
    cells = []
    for it in text_items:
        if it[5] != table[0].get("page"):
            continue
        if it[7] <= baseline + 8 or not (x_min <= it[6] <= x_max):
            continue
        col = min(range(len(xs)), key=lambda i: abs(it[6] - xs[i]))
        cells.append((col, it[7], it[6], it[1].strip()))
    cells.sort(key=lambda c: (c[1], c[0]))
    grouped, last_y = [], None
    for cell in cells:
        if last_y is None or cell[1] - last_y <= 12:
            if not grouped or (last_y is not None and cell[1] - last_y > 12):
                grouped.append([])
            grouped[-1].append(cell)
            last_y = cell[1] if last_y is None else max(last_y, cell[1])
        elif cell[1] - last_y > 48:
            break
        else:
            grouped.append([cell])
            last_y = cell[1]
    records = []
    for group in grouped:
        buckets = {"tag": [], "model": [], "serial": [], "part": [], "ref": [], "related": []}
        positions = []
        for col, y, x, text in group:
            if not text or _field_of_label(text):
                continue
            role = _cell_role(table[col]["field"], text)
            if role == "serial":
                pair = (table[col].get("label", ""), text)
                if pair not in buckets["serial"]:
                    buckets["serial"].append(pair)
            else:
                _append_unique(buckets[role], text)
            positions.append((norm(text), x, y))
        if not buckets["tag"] and not (buckets["serial"] and buckets["ref"]):
            continue
        tags = sorted(buckets["tag"], key=len, reverse=True)
        primary = tags[0] if tags else ""
        for extra in tags[1:]:
            _append_unique(buckets["related"], extra)
        if not buckets["model"]:
            codes = [v for v in buckets["related"]
                     if len(v) >= 10 and "-" not in v and any(c.isalpha() for c in v) and any(c.isdigit() for c in v)]
            if codes:
                chosen = max(codes, key=len)
                buckets["model"] = [chosen]
                buckets["related"] = [v for v in buckets["related"] if v != chosen]
        row_y = sum(c[1] for c in group) / len(group)
        serial = _pick_serial(buckets["serial"])
        records.append({
            "page": table[0]["page"],
            "y": row_y,
            "tag": primary,
            "model": _normalize_model(" | ".join(buckets["model"])),
            "serial": serial,
            "part": " | ".join(buckets["part"]),
            "ref": " | ".join(buckets["ref"]),
            "related": " | ".join(buckets["related"]),
            "cells": positions,
        })
    return records


def _add_extra_headers(headers, text_items):
    """Keep unnamed columns that sit on the same heading line, such as a line class."""
    if not headers:
        return headers
    bands, extras = [], []
    for header in sorted(headers, key=lambda h: h["y"]):
        if bands and abs(header["y"] - bands[-1][0]["y"]) <= 18:
            bands[-1].append(header)
        else:
            bands.append([header])
    occupied = {(round(h["x"]), round(h["y"])) for h in headers}
    for band in bands:
        y0 = min(h["y"] for h in band) - 14
        y1 = max(h["y"] for h in band) + 8
        xs = sorted(h["x"] for h in band)
        for it in text_items:
            label = it[1].strip()
            if not label or len(label) > 36:
                continue
            if not (y0 <= it[7] <= y1) or it[6] < xs[0] - 30 or it[6] > xs[-1] + 170:
                continue
            spot = (round(it[6]), round(it[7]))
            if spot in occupied:
                continue
            if any(abs(it[6] - h["x"]) < 40 and abs(it[7] - h["y"]) < 16 for h in headers):
                continue
            extras.append({
                "x": it[6], "y": it[7], "field": _field_of_label(label) or "extra",
                "label": label, "page": it[5],
            })
            occupied.add(spot)
    return headers + extras


def index_table_records(items):
    """Rows from tag / serial / P&ID lists, keyed so a hit can find its row."""
    by_page = {}
    for it in _page_text_items(items):
        by_page.setdefault(it[5], []).append(it)
    index = {}
    for page_items in by_page.values():
        headers, _used = _header_cells(page_items)
        for header in headers:
            header["page"] = page_items[0][5]
        headers = _add_extra_headers(headers, page_items)
        for table in _tables_from_headers(headers):
            for rec in _records_from_table(table, page_items):
                for text_key, x, y in rec["cells"]:
                    index.setdefault((rec["page"], text_key), []).append((x, y, rec))
    return index


def lookup_record(index, item):
    best, best_d = None, 22
    for x, y, rec in index.get((item[5], norm(item[1])), ()):
        dist = ((x - item[6]) ** 2 + (y - item[7]) ** 2) ** 0.5
        if dist < best_d:
            best, best_d = rec, dist
    return best


# Words that mean "list every tag row", not "find this text".
_LIST_QUERIES = {
    "TAG", "TAGNO", "TAGS", "TAGNUMBER",
    "SERIAL", "SERIALNO", "SERIALNUMBER", "SRNO", "SRNUMBER",
    "PART", "PARTNO", "PARTNUMBER",
    "MODEL", "MODELNO", "ITEMCODE", "ITEMNO",
}


def is_list_query(q):
    return norm(q) in _LIST_QUERIES


def unique_records(index):
    seen, out = set(), []
    for hits in index.values():
        for _x, _y, rec in hits:
            key = id(rec)
            if key in seen:
                continue
            seen.add(key)
            out.append(rec)
    return out


def _blank_row(query, page_no, page_meta):
    return {
        "Query": query, "Tag / Value": "", "Type": "Tag list", "Function": "",
        "Function Description": "", "Loop Number": "", "Suffix": "", "Page": page_no,
        "Drawing Number": page_meta.get("dwg", ""), "Drawing Title": page_meta.get("title", ""),
        "Revision": page_meta.get("rev", ""), "X (pt)": "", "Y (pt)": "",
        "Nearby Text": "", "Full Line Text": "", "Source": "list",
    }


def _serial_sort(value):
    value = (value or "").strip()
    return (0, int(value)) if value.isdigit() else (1, value)


def _listed_tag(tag):
    """Keep a real tag number and drop headings or broken OCR fragments."""
    tag = (tag or "").strip()
    if (not tag) or (" " in tag) or not any(ch.isdigit() for ch in tag):
        return ""
    found = re.search(
        r"(?:\d{3}-){1,2}[A-Z]{1,5}-\d{3,6}[A-Z]{0,3}|[A-Z]{1,5}-\d{3,6}[A-Z]{0,3}",
        tag, re.I)
    if found and len(tag) - len(found.group(0)) <= 2:
        return found.group(0)
    return tag if len(tag) >= 6 else ""


def rows_from_records(query, records, pages):
    """One result per tag row, with that row's own serial, model, part, and reference."""
    emitted, seen = [], set()
    ordered = sorted(records, key=lambda r: (r["page"], _serial_sort(r["serial"]), r["tag"]))
    for rec in ordered:
        tag = _listed_tag(rec.get("tag", ""))
        if not tag:
            continue
        rec = dict(rec)
        rec["tag"] = tag
        sig = (rec["page"], rec["tag"], rec["serial"], rec["model"], rec["part"], rec["ref"])
        if sig in seen:
            continue
        seen.add(sig)
        pm = pages.get(rec["page"], {})
        row = _blank_row(query, rec["page"], pm)
        row["Tag / Value"] = rec["tag"] or rec["serial"]
        row["Y (pt)"] = round(rec["y"], 1)
        fill_related(row, rec, pm)
        emitted.append(row)
    return emitted


def same_tag(query, text, rec):
    """True when this hit is the tag asked for, not a longer line that merely contains it."""
    nq = norm(query)
    for value in (text, rec["tag"] if rec else ""):
        nt = norm(value or "")
        if not nt:
            continue
        if nt == nq:
            return True
        extra = nt[len(nq):] if nt.startswith(nq) else ""
        if extra.isalpha() and 1 <= len(extra) <= 3:
            return True
    return False


def fill_related(row, rec, page_meta):
    """Fill tag, model/item code, serial, part number, and reference file."""
    drawing = page_meta.get("dwg", "")
    if rec:
        row["Tag Number"] = rec["tag"]
        row["Model or Item Code"] = rec["model"]
        row["Serial No"] = rec["serial"]
        row["Part Number"] = rec["part"]
        row["Reference File"] = rec["ref"] or drawing
        row["Related Values"] = rec["related"]
        return
    text = row["Tag / Value"] or ""
    kind = row["Type"]
    if kind in ("Instrument", "Equipment", "Line No.") or RE_ROW_TAG.match(text):
        row["Tag Number"] = text
    else:
        row["Tag Number"] = ""
    if kind == "Line No." or (kind != "Instrument" and (RE_ITEM_CODE.search(text) or RE_LINE_NO.search(text))):
        row["Model or Item Code"] = text
    else:
        row["Model or Item Code"] = ""
    row["Serial No"] = ""
    row["Part Number"] = ""
    row["Reference File"] = drawing
    row["Related Values"] = ""


def search(data, queries, nearby_radius=90, nearby_count=6):
    """Search ONE analysed PDF. Returns (rows, summary) - rows have no file name yet."""
    pages = {p["page"]: p for p in data["pages"]}
    items = data["items"]
    by_page = {}
    for it in items:
        by_page.setdefault(it[5], []).append(it)
    record_index = index_table_records(items)

    rows, summary = [], []
    for q in queries:
        if is_list_query(q):
            emitted = rows_from_records(q, unique_records(record_index), pages)
            rows.extend(emitted)
            summary.append({"Query": q, "Matches": len(emitted),
                            "Pages": sorted({h["Page"] for h in emitted})})
            continue
        match = build_matcher(q)
        seen, hits = set(), []
        func_query = norm(q).isalpha() and len(norm(q)) <= 6 and not q.lower().startswith("re:")
        for it in items:
            if func_query and it[0] == "Text":
                continue
            if match(it):
                key = (it[0], it[1], it[5], it[6], it[7])
                if key in seen:
                    continue
                seen.add(key)
                hits.append(it)
        if func_query and not hits:        # not a function code -> plain word search
            hits = [it for it in items if norm(it[1]) == norm(q)]
        # a Text hit on the same spot as an Instrument/Line/Equipment hit is a duplicate
        special = {(h[5], h[6], h[7]) for h in hits if h[0] != "Text"}
        hits = [h for h in hits if h[0] != "Text" or (h[5], h[6], h[7]) not in special]
        hits.sort(key=lambda h: (h[5], h[7], h[6]))

        emitted, seen_rows = [], set()
        broad = not func_query and not q.lower().startswith("re:") and "*" not in q and "?" not in q
        for it in hits:
            kind, text, func, num, suf, pno, x, y, src, full = it
            rec = lookup_record(record_index, it)
            if broad and not same_tag(q, text, rec):
                continue
            if rec is not None:
                rid = (pno, round(rec["y"], 1), rec["tag"], rec["serial"], rec["ref"])
                if rid in seen_rows:
                    continue
                seen_rows.add(rid)
            near = []
            for o in by_page.get(pno, []):
                if o[0] == "Text" and o is not it:
                    d = ((o[6] - x) ** 2 + (o[7] - y) ** 2) ** 0.5
                    if 0 < d <= nearby_radius and norm(o[1]) != norm(text):
                        near.append((d, o[1]))
            near.sort()
            nb = " | ".join(t for _, t in near[:nearby_count])
            pm = pages.get(pno, {})
            row = {"Query": q, "Tag / Value": text, "Type": kind, "Function": func,
                   "Function Description": decode_function(func) if kind == "Instrument" else "",
                   "Loop Number": num, "Suffix": suf, "Page": pno,
                   "Drawing Number": pm.get("dwg", ""), "Drawing Title": pm.get("title", ""),
                   "Revision": pm.get("rev", ""), "X (pt)": x, "Y (pt)": y,
                   "Nearby Text": nb, "Full Line Text": full if kind != "Text" or full != text else text,
                   "Source": src}
            fill_related(row, rec, pm)
            emitted.append(row)
        rows.extend(emitted)
        summary.append({"Query": q, "Matches": len(emitted),
                        "Pages": sorted({h["Page"] for h in emitted})})
    return rows, summary


def pages_str(pg, limit=60):
    pg = sorted(pg)
    return ", ".join(str(p) for p in pg[:limit]) + (" ..." if len(pg) > limit else "")


def search_all(datasets, queries):
    """Search every analysed file. Returns (rows, summary_rows, per_query_totals)."""
    valid = []
    for q in queries:
        try:
            build_matcher(q)
            valid.append(q)
        except re.error as e:
            print(f"  [!] Bad regular expression '{q}': {e}")

    rows, per_file = [], {q: [] for q in valid}
    for display, path, data in datasets:
        r, s = search(data, valid)
        for row in r:
            row["File Name"] = display
            row["File Path"] = path
        rows.extend(r)
        for item in s:
            per_file[item["Query"]].append((display, item))

    order = {q: i for i, q in enumerate(valid)}
    rows.sort(key=lambda r: order[r["Query"]])      # stable: keeps file order, then page order

    summary, totals = [], []
    for q in valid:
        tot_m, files_hit = 0, 0
        for display, item in per_file[q]:
            if item["Matches"]:
                files_hit += 1
                tot_m += item["Matches"]
                summary.append({"Query": q, "File Name": display, "Matches": item["Matches"],
                                "Pages": len(item["Pages"]), "Status": "FOUND",
                                "Page list": pages_str(item["Pages"])})
        if not tot_m:
            summary.append({"Query": q, "File Name": "-", "Matches": 0, "Pages": 0,
                            "Status": "NOT FOUND", "Page list": ""})
        totals.append({"Query": q, "Matches": tot_m, "Files": files_hit,
                       "Status": "FOUND" if tot_m else "NOT FOUND"})
    return rows, summary, totals


# --------------------------------------------------------------------------
#  Excel
# --------------------------------------------------------------------------
def write_excel(path, rows, summary, columns, source_name, datasets):
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    head_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(bold=True, color="FFFFFF")

    ws.append(columns)
    for r in rows:
        ws.append([r.get(c, "") for c in columns])
    for c in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = head_fill, head_font
        cell.alignment = Alignment(vertical="center")
        width = max([len(str(columns[c - 1]))] + [len(str(ws.cell(row=r, column=c).value or ""))
                                                  for r in range(2, min(ws.max_row, 400) + 1)])
        ws.column_dimensions[get_column_letter(c)].width = min(max(10, width + 2), 70)
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions

    # ---- Summary: per query, per file ---------------------------------------
    s2 = wb.create_sheet("Summary")
    s2.append(["Source", source_name])
    s2.append(["Files analysed", len(datasets)])
    s2.append(["Created", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    s2.append([])
    hdr = ["Query", "File Name", "Matches", "Pages", "Status", "Page list"]
    s2.append(hdr)
    for c in range(1, len(hdr) + 1):
        s2.cell(row=5, column=c).fill, s2.cell(row=5, column=c).font = head_fill, head_font
    for s in summary:
        s2.append([s["Query"], s["File Name"], s["Matches"], s["Pages"], s["Status"], s["Page list"]])
    for col, w in zip("ABCDEF", (28, 45, 10, 8, 12, 80)):
        s2.column_dimensions[col].width = w
    s2.freeze_panes = "A6"

    # ---- Unique tags with file + page locations -----------------------------
    s3 = wb.create_sheet("Unique Tags")
    uh = ["Tag / Value", "Type", "Occurrences", "Files", "Locations (file: pages)"]
    s3.append(uh)
    for c in range(1, len(uh) + 1):
        s3.cell(row=1, column=c).fill, s3.cell(row=1, column=c).font = head_fill, head_font
    uniq = {}
    for r in rows:
        u = uniq.setdefault((r["Tag / Value"], r["Type"]), {"n": 0, "f": {}})
        u["n"] += 1
        u["f"].setdefault(r["File Name"], set()).add(r["Page"])
    for (tag, typ), u in sorted(uniq.items()):
        loc = " ; ".join(f"{f}: {pages_str(p, 30)}" for f, p in u["f"].items())
        s3.append([tag, typ, u["n"], len(u["f"]), loc[:32000]])
    for col, w in zip("ABCDE", (34, 12, 12, 8, 100)):
        s3.column_dimensions[col].width = w
    s3.freeze_panes = "A2"
    s3.auto_filter.ref = s3.dimensions

    # ---- Files analysed ------------------------------------------------------
    s4 = wb.create_sheet("Files")
    fh = ["File Name", "Pages", "Indexed items", "Full path"]
    s4.append(fh)
    for c in range(1, len(fh) + 1):
        s4.cell(row=1, column=c).fill, s4.cell(row=1, column=c).font = head_fill, head_font
    for display, p, data in datasets:
        s4.append([display, len(data["pages"]), len(data["items"]), p])
    for col, w in zip("ABCD", (45, 8, 14, 90)):
        s4.column_dimensions[col].width = w
    wb.save(path)


# --------------------------------------------------------------------------
#  Interactive flow
# --------------------------------------------------------------------------
def split_csv(s):
    return [x.strip() for x in re.split(r",", s) if x.strip()]


def resolve_columns(sel):
    """Turn column names or numbers into an ordered column list.

    Returns (columns, unknown). An empty selection uses the default set.
    The single value "all" selects every column. File Name and Page are
    always included so every hit can be located.
    """
    if not sel:
        return list(DEFAULT_COLUMNS), []
    if len(sel) == 1 and str(sel[0]).strip().lower() == "all":
        return list(COLUMNS), []
    out, unknown = [], []
    lower = {c.lower(): c for c in COLUMNS}
    for raw in sel:
        s = str(raw).strip()
        if not s:
            continue
        if s.isdigit() and 1 <= int(s) <= len(COLUMNS):
            c = COLUMNS[int(s) - 1]
        elif s.lower() in lower:
            c = lower[s.lower()]
        else:
            cands = [c for c in COLUMNS if s.lower() in c.lower()]
            if len(cands) == 1:
                c = cands[0]
            else:
                unknown.append(s)
                continue
        if c not in out:
            out.append(c)
    for must in REQUIRED_COLUMNS:
        if must not in out:
            out.append(must)
    return (out or list(DEFAULT_COLUMNS)), unknown


def ask_columns(preset=None):
    if preset:
        sel = split_csv(preset)
    else:
        print("\nAvailable columns:")
        for i, c in enumerate(COLUMNS, 1):
            print(f"  {i:2}. {c}")
        print("\nWhich columns do you want in the Excel sheet?")
        print("  Type names or numbers separated by commas  (e.g.  Tag / Value, File Name, Page)")
        print("  Press Enter = default (%s)   |   type 'all' for everything" % ", ".join(DEFAULT_COLUMNS))
        sel = split_csv(input("Columns> ").strip())
    cols, unknown = resolve_columns(sel)
    for s in unknown:
        print(f"  [!] Unknown column '{s}' - ignored")
    return cols


def in_colab():
    return "google.colab" in sys.modules


def find_default_inputs():
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:                        # notebooks have no __file__
        here = os.getcwd()
    for folder in (os.getcwd(), here):
        zips = [f for f in os.listdir(folder) if f.lower().endswith(".zip")]
        pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        if len(zips) == 1 and not pdfs:
            return [os.path.join(folder, zips[0])]
        if pdfs and not zips:
            return [os.path.join(folder, f) for f in pdfs] if len(pdfs) > 1 else [os.path.join(folder, pdfs[0])]
    return []


def main():
    ap = argparse.ArgumentParser(description="Extract tags/values from PDF/P&ID files (single, ZIP or folder) into Excel")
    ap.add_argument("inputs", nargs="*", help="PDF file(s), ZIP file(s) and/or folder(s)")
    ap.add_argument("--tags", help="comma separated tags (skips the question)")
    ap.add_argument("--columns", help="comma separated columns (skips the question)")
    ap.add_argument("--out", help="output .xlsx file")
    ap.add_argument("--ocr", choices=["auto", "off", "all"], default="auto",
                    help="OCR for scanned/image pages (default auto)")
    ap.add_argument("--no-cache", action="store_true", help="re-analyse even if a cache exists")
    ap.add_argument("--web", action="store_true", help="open the browser app")
    a, _unknown = ap.parse_known_args()      # ignores Jupyter/Colab's own "-f kernel.json"
    if a.web:
        from app import serve
        serve()
        return
    inputs = [p for p in a.inputs if p.lower().endswith((".pdf", ".zip")) or os.path.isdir(p)]

    print("=" * 66)
    print("   PDF / P&ID TAG EXTRACTOR  ->  EXCEL   (multi-file / ZIP)")
    print("=" * 66)

    if not inputs:
        d = find_default_inputs()
        if not d and in_colab():
            from google.colab import files
            print("Select your PDF(s) or ZIP to upload...")
            up = files.upload()
            d = [k for k in up if k.lower().endswith((".pdf", ".zip"))]
        if d and in_colab():
            inputs = d
        else:
            shown = (" [Enter = %s]" % (os.path.basename(d[0]) if len(d) == 1 else f"{len(d)} PDFs here")) if d else ""
            raw = input(f"PDF / ZIP / folder path{shown} (several allowed, separated by commas): ").strip()
            inputs = [x.strip().strip('"') for x in raw.split(",") if x.strip()] or d

    files = collect_pdfs(inputs)
    if not files:
        sys.exit("No PDF files found in: %r" % inputs)
    print(f"\nFound {len(files)} PDF file(s):")
    for display, _p in files[:15]:
        print("   -", display)
    if len(files) > 15:
        print(f"   ... and {len(files) - 15} more")

    datasets = analyze_all(files, a.ocr, use_cache=not a.no_cache)
    if not datasets:
        sys.exit("None of the PDFs could be analysed.")

    kinds, total_pages = {}, 0
    for _d, _p, data in datasets:
        total_pages += len(data["pages"])
        for it in data["items"]:
            kinds[it[0]] = kinds.get(it[0], 0) + 1
    print(f"\n  Done: {len(datasets)} file(s), {total_pages} pages  |  " +
          "  |  ".join(f"{k}: {v:,}" for k, v in sorted(kinds.items())))

    base_dir = os.path.dirname(os.path.abspath(inputs[0])) if not os.path.isdir(inputs[0]) else os.path.abspath(inputs[0])
    source_name = os.path.basename(inputs[0]) if len(inputs) == 1 else f"{len(inputs)} inputs"
    if len(inputs) == 1 and os.path.isdir(inputs[0]):
        source_name = os.path.basename(os.path.abspath(inputs[0])) + " (folder)"

    first_round = True
    while True:
        if a.tags and first_round:
            tags_in = a.tags
        else:
            print("\nWhich tag(s) / value(s) do you want to extract (searched in ALL files)?")
            print("  Several allowed, separated by commas.  Examples:")
            print("    PIT-01043, TIT, 866-E-3303B, *3CA1P01*, 01043")
            print("  (type 'help' for all search options, 'q' to quit)")
            tags_in = input("Tags> ").strip()
            if tags_in.lower() in ("q", "quit", "exit"):
                break
            if tags_in.lower() == "help":
                print(__doc__)
                continue
        queries = split_csv(tags_in)
        if not queries:
            continue

        rows, summary, totals = search_all(datasets, queries)
        print("\n  Results:")
        for t in totals:
            print(f"    {t['Query']:<28} {t['Matches']:>6} matches in {t['Files']:>3} file(s)   {t['Status']}")
        if not rows:
            print("  Nothing found - try another spelling or a wildcard like *PIT*")
            if a.tags:
                break
            continue

        columns = ask_columns(a.columns if first_round else None)
        out = a.out if (a.out and first_round) else os.path.join(
            base_dir, "extracted_tags_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S"))
        write_excel(out, rows, summary, columns, source_name, datasets)
        if in_colab():
            try:
                from google.colab import files as colab_files
                colab_files.download(out)
            except Exception:
                pass
        print(f"\n  Excel written: {out}\n  Rows: {len(rows):,}   Columns: {', '.join(columns)}")

        first_round = False
        if a.tags:
            break
        again = input("\nSearch for more tags? (y/n) [y]: ").strip().lower()
        if again in ("n", "no"):
            break
    print("\nBye.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
