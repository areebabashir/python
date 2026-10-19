#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP EQUIPMENT DATA EXTRACTOR  (vendor documents / MDR / test certificates -> Excel)
==================================================================================

Reads one PDF, many PDFs, or whole folders of PDFs and pulls out, per TAG:

    Tag Number | Equipment (VALVE / ACTUATOR / ...) | Manufacturer | Model Number
    | Serial Number | Part Number | Item Code | Description

and records WHERE every value came from (file name + page number + how it was read).

Works on both kinds of PDF:
  * "text" PDFs  (copyable text)      -> read directly, 100 % accurate characters
  * scanned PDFs (images of pages)    -> OCR with Tesseract, incl. tables and
                                         sideways / upside-down pages

How it stays accurate on noisy scans
  * Tables are cut into individual cells and each important cell is read several
    times with different image settings; the readings vote.
  * Every value is collected from EVERY page it appears on, then the values vote
    across all pages / all files (e.g. the serial appears on the test report, the
    test certificate, the PMI report ... -> majority wins).
  * Typical OCR confusions (7<->1, C<->0, O<->0, S<->5 ...) are repaired using the
    tag and serial-number formats learned from the batch itself.
  * Optional: give a master tag list (--taglist) and every OCR tag is snapped to
    the nearest real tag. STRONGLY recommended for scanned documents.
  * Anything uncertain is flagged in the "Review" column - check those rows.

USAGE
-----
  python sap_equipment_extractor.py

  A window opens every time:
    1. "Add PDF file(s)..." or "Add a whole folder..."  -> upload your PDFs (one or many)
    2. optional: "Upload tag list..." (Excel/CSV with your real tag numbers - improves OCR accuracy)
       reading mode Fast / Normal / Thorough, and where to save the Excel result
    3. "START EXTRACTION"  -> progress bar, then the Excel result opens

  Advanced (no window):  python sap_equipment_extractor.py --no-window "D:\\MDR" --recursive -o result.xlsx

INSTALL (Windows)
-----------------
  1. Install Tesseract OCR:  https://github.com/UB-Mannheim/tesseract/wiki
     (default path C:\\Program Files\\Tesseract-OCR\\tesseract.exe is found automatically)
  2. pip install pymupdf pytesseract opencv-python-headless numpy pillow openpyxl
"""

import os
import re
import sys
import glob
import argparse
import datetime
import difflib
import traceback
from collections import defaultdict, Counter
from dataclasses import dataclass, field, asdict
from concurrent.futures import ProcessPoolExecutor, as_completed

import shutil
import subprocess
import threading


def _in_colab():
    try:
        import google.colab  # noqa: F401
        return True
    except ImportError:
        return False


IN_COLAB = _in_colab()


def _bootstrap():
    """Install missing Python packages (and Tesseract on Colab) automatically."""
    need = []
    for mod, pkg in (("pytesseract", "pytesseract"), ("cv2", "opencv-python-headless"),
                     ("openpyxl", "openpyxl"), ("numpy", "numpy"), ("PIL", "pillow")):
        try:
            __import__(mod)
        except ImportError:
            need.append(pkg)
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        try:
            import fitz  # noqa: F401
        except ImportError:
            need.append("pymupdf")
    if need:
        print("Installing:", " ".join(need), "...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *need])
    if IN_COLAB and shutil.which("tesseract") is None:
        print("Installing Tesseract OCR (one time, ~1 minute) ...")
        subprocess.run("apt-get -qq update > /dev/null && apt-get -qq install -y tesseract-ocr > /dev/null",
                       shell=True, check=False)


_bootstrap()

try:
    import pymupdf  # PyMuPDF >= 1.24
except ImportError:  # older PyMuPDF
    import fitz as pymupdf

os.environ.setdefault("OMP_THREAD_LIMIT", "1")   # one thread per tesseract call; we parallelise pages instead

import numpy as np
import cv2
import pytesseract


# =============================================================================
#  CONFIGURATION  -  adjust these to your documents if needed
# =============================================================================

# Name used for the main equipment when a value is not specifically about the
# actuator / positioner / etc.
DEFAULT_COMPONENT = "VALVE"

# Words in a line (or in a page header) that mean the value belongs to another
# component of the same tag.  First match wins.
COMPONENT_KEYWORDS = [
    ("ACTUATOR",       [r"\bactuator\b", r"\bact\.?\s*s[il1]\.?\s*no", r"\bact\.?\s*serial"]),
    ("POSITIONER",     [r"\bpositioner\b"]),
    ("SOLENOID VALVE", [r"\bsolenoid\b", r"\bsov\b"]),
    ("LIMIT SWITCH",   [r"\blimit\s*switch", r"\bl/?s\s*box\b"]),
    ("GEARBOX",        [r"\bgear\s*box\b", r"\bgearbox\b"]),
]

# Known manufacturers -> canonical name written to Excel.
# Keys are matched case-insensitively inside the text.  Add your vendors here.
KNOWN_MANUFACTURERS = {
    "l&t valves": "L&T Valves Limited",
    "l & t valves": "L&T Valves Limited",
    "larsen & toubro valves": "L&T Valves Limited",
    "rotork": "Rotork",
    "emerson": "Emerson", "fisher": "Fisher (Emerson)", "bettis": "Bettis (Emerson)",
    "keystone": "Keystone (Emerson)", "el-o-matic": "El-O-Matic (Emerson)",
    "rosemount": "Rosemount (Emerson)", "flowserve": "Flowserve", "limitorque": "Limitorque (Flowserve)",
    "valtek": "Valtek (Flowserve)", "cameron": "Cameron", "metso": "Metso", "neles": "Neles",
    "valmet": "Valmet", "samson": "Samson", "auma": "AUMA", "biffi": "Biffi",
    "bray": "Bray", "velan": "Velan", "ksb": "KSB", "asco": "ASCO", "abb": "ABB",
    "siemens": "Siemens", "yokogawa": "Yokogawa", "honeywell": "Honeywell",
    "endress": "Endress+Hauser", "krohne": "Krohne", "vega": "VEGA", "wika": "WIKA",
    "baker hughes": "Baker Hughes", "masoneilan": "Masoneilan", "habonim": "Habonim",
    "petrolvalves": "PetrolValves", "pibiviesse": "Pibiviesse", "pbv": "PBV",
    "virgo": "Virgo Valves", "audco": "Audco", "kitz": "KITZ", "tyco": "Tyco",
    "westlock": "Westlock", "topworx": "TopWorx", "norgren": "Norgren", "parker": "Parker",
    "swagelok": "Swagelok", "circor": "CIRCOR", "mokveld": "Mokveld", "oliver valves": "Oliver Valves",
    "schuck": "Schuck", "orbit": "Orbit", "copes-vulcan": "Copes-Vulcan", "imi ": "IMI",
}

# Manufacturers that make actuators: a page whose header shows one of these is
# treated as an actuator document.
ACTUATOR_BRANDS = {"Rotork", "AUMA", "Biffi", "Bettis (Emerson)", "Limitorque (Flowserve)",
                   "El-O-Matic (Emerson)"}

# Extra tag formats (Python regex) if yours are not like 775-730-AOV-0724.
# Use ONE capture group for the whole tag.  The built-in pattern already covers
#   NNN-NNN-AAA-NNNN  (2-4 digits, 2-4 digits, 1-6 letters, 2-5 digits, opt. suffix letter)
EXTRA_TAG_PATTERNS = [
    # r"\b(\d{2,3}-[A-Z]{1,5}-\d{3,5}[A-Z]?)\b",        # e.g. 30-PSV-1001A
]

# Serial number formats you already know (letters literal, '9' = any digit).
# Formats are ALSO learned automatically from the documents.
SERIAL_FORMATS = [
    # "99C999999",      # L&T Valves  e.g. 25C123258
    # "BR999999999",    # Rotork      e.g. BR758900701
]

# Words on the same line that mean a "serial / model" is NOT the equipment's
# (test instruments, gauges, heat numbers ...)
NEGATIVE_CONTEXT = [
    "gauge", "calibration", "heat", "batch", "coupon", "machine", "instrument",
    "meter", "thermometer", "stop watch", "vernier", "mtc", "report no", "pressure gauge",
    "penetrant", "developer", "cleaner", "invoice", "challan", "dc no", "lot no",
]

# Pages that are sub-supplier material certificates (forgings, bolts, gaskets ...)
# -> never used as a source for equipment-level data.
COMPONENT_CERT_KEYWORDS = [
    "material test certificate", "mill test certificate", "material certificate",
    "inspection certificate (", "chemical composition", "chemical properties",
    "chemical analysis",
]

# =============================================================================
#  END OF CONFIGURATION
# =============================================================================

FIELDS = ["MANUFACTURER", "MODEL", "SERIAL", "PART", "ITEM_CODE", "TYPE", "DESCRIPTION"]
FIELD_TITLES = {
    "MANUFACTURER": "Manufacturer", "MODEL": "Model Number", "SERIAL": "Serial Number",
    "PART": "Part Number", "ITEM_CODE": "Item Code", "TYPE": "Type", "DESCRIPTION": "Description",
}

# -----------------------------------------------------------------------------
#  text helpers
# -----------------------------------------------------------------------------
_DASHES = "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"
_QUOTES = {"\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'", "\u2033": '"'}


def clean_text(s: str) -> str:
    if not s:
        return ""
    for d in _DASHES:
        s = s.replace(d, "-")
    for k, v in _QUOTES.items():
        s = s.replace(k, v)
    return s.replace("\t", " ").strip()


_DIGIT_FIX = {"O": "0", "o": "0", "Q": "0", "D": "0", "I": "1", "l": "1", "|": "1",
              "i": "1", "S": "5", "s": "5", "Z": "2", "z": "2", "B": "8"}
_LETTER_FIX = {"0": "O", "1": "I", "5": "S", "8": "B", "2": "Z", "6": "G"}

# characters OCR typically swaps for each other (both directions)
_CONFUSABLE = set()
for a, b in ["17", "71", "08", "80", "38", "83", "56", "65", "68", "86", "59", "95",
             "0O", "O0", "OQ", "QO", "OD", "DO", "C0", "0C", "CG", "GC", "C6", "6C",
             "1I", "I1", "1L", "L1", "5S", "S5", "2Z", "Z2", "8B", "B8", "49", "94",
             "47", "74", "27", "72", "13", "31"]:
    _CONFUSABLE.add((a, b))


def fix_digits(seg: str) -> str:
    return "".join(_DIGIT_FIX.get(c, c) for c in seg)


def fix_letters(seg: str) -> str:
    return "".join(_LETTER_FIX.get(c, c) for c in seg.upper())


_DIG = r"[0-9OoQDIl|SZB]"
TAG_RX = re.compile(
    rf"(?<![A-Za-z0-9,\-/])({_DIG}{{2,4}})\s*[-_.;:~]\s*({_DIG}{{2,4}})\s*[-_.;:~]\s*"
    rf"([A-Za-z0Q]{{1,6}})\s*(?:[-_.;:~]\s*|\s)({_DIG}{{2,5}})([A-Z])?(?![A-Za-z0-9])"
)
EXTRA_TAG_RX = [re.compile(p) for p in EXTRA_TAG_PATTERNS]


def find_tags(text: str):
    """Return list of (normalized_tag, raw_text) found in text."""
    out = []
    text = clean_text(text)
    for m in TAG_RX.finditer(text):
        s1, s2, let, num, suf = m.groups()
        ok = True
        for seg in (s1, s2, num):
            if sum(c.isdigit() for c in seg) < len(seg) - 1:
                ok = False
        if sum(c.isalpha() for c in let) < max(1, len(let) - 1):
            ok = False
        if not ok:
            continue
        tag = f"{fix_digits(s1)}-{fix_digits(s2)}-{fix_letters(let)}-{fix_digits(num)}{suf or ''}"
        out.append((tag, m.group(0)))
    for rx in EXTRA_TAG_RX:
        for m in rx.finditer(text):
            out.append((m.group(1).upper(), m.group(0)))
    return out


DATE_RX = re.compile(r"\d{1,4}[./-]\d{1,2}[./-]\d{1,4}")


def code_like(tok: str, min_digits=3, min_len=5) -> bool:
    t = tok.strip(".,:;|()[]{}'\"_-~*")
    if not (min_len <= len(t) <= 30):
        return False
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-/_.#]*", t):
        return False
    if DATE_RX.fullmatch(t):
        return False
    return sum(c.isdigit() for c in t) >= min_digits


def strip_junk(s: str) -> str:
    return s.strip(" .,:;|()[]{}'\"_-~*/\\<>=+`^!?°©®«»¢€£$%&#@").strip()


def first_code_token(text: str, min_digits=3, min_len=5, allow_join=True):
    """First serial/part-like token in text. Joins '25C 123258' -> '25C123258'."""
    toks = [strip_junk(t) for t in re.split(r"\s+", text) if strip_junk(t)]
    for i, t in enumerate(toks):
        if allow_join and i + 1 < len(toks):
            j = t + toks[i + 1]
            if (len(t) <= 4 and re.fullmatch(r"[A-Za-z0-9]+", toks[i + 1])
                    and code_like(j, min_digits, min_len) and not code_like(t, min_digits, min_len)
                    and not re.fullmatch(r"\d{1,2}", t)):      # '1 25C123258' = qty + serial, don't glue
                return j
        if code_like(t, min_digits, min_len):
            return t
    return None


def shape_of(v: str) -> str:
    return "".join("9" if c.isdigit() else c.upper() for c in v)


def canon_manufacturer(text: str, fuzzy=False):
    """Known manufacturer mentioned in text (earliest mention wins). fuzzy=True also
    accepts OCR-damaged spellings such as 'C&T VALVES LIMITED' or 'rotori'."""
    low = " " + re.sub(r"\s*&\s*", "&", text.lower()) + " "
    best, pos = None, 10 ** 9
    for key, name in KNOWN_MANUFACTURERS.items():
        k = re.sub(r"\s*&\s*", "&", key.lower().strip())
        m = re.search(r"(?<![a-z])" + re.escape(k) + r"(?![a-z])", low)
        if m and m.start() < pos:
            best, pos = name, m.start()
    if best or not fuzzy:
        return best
    words = re.findall(r"[a-z&]+", low)
    for key, name in KNOWN_MANUFACTURERS.items():
        k = re.sub(r"\s*&\s*", "&", key.lower().strip())
        kw = k.split()
        if len(k) < 5:
            continue
        for i in range(len(words) - len(kw) + 1):
            cand = " ".join(words[i:i + len(kw)])
            if difflib.SequenceMatcher(None, cand, k).ratio() >= 0.8:
                return name
    return None


def component_of(line_lower: str):
    for comp, pats in COMPONENT_KEYWORDS:
        for p in pats:
            if re.search(p, line_lower):
                return comp
    return None


# -----------------------------------------------------------------------------
#  label based extraction (key : value lines)
# -----------------------------------------------------------------------------
_STOP_WORDS = (r"\s(?:vtr|ref|tc\s*no|>|part|p/n|tag|tag\s*no|job|item|model|serial|s/n|date|customer|project|po|qty|"
               r"quantity|order|page|rev|size|rating|class|group|end\s*user|catalogue|catalog|"
               r"code|actuator|make|type|sl\s*no|external|so\s*item|tc\s*ref)\b")

LABELS = [
    # (field, regex for the label, strength 1-3)
    ("SERIAL", r"(?:valve|act(?:uator)?\.?|equipment|unit|mfr\.?|manufacturer'?s?)\s*"
               r"(?:serial\s*(?:no|num(?:ber)?s?|#)?|s[il1]\.?\s*no\.?s?|sr\.?\s*no\.?|s/n)", 3),
    ("SERIAL", r"serial\s*(?:no|num(?:ber)?s?|#)?\.?", 2),
    ("SERIAL", r"\bs/n\b", 2),
    ("MODEL", r"(?:actuator\s*)?model\s*(?:no\.?|number|#)?", 3),
    ("MODEL", r"(?:figure|fig\.?)\s*(?:no\.?|number)", 3),
    ("MODEL", r"catalog(?:ue)?\b(?!\s*&)\s*(?:no\.?|number)?", 2),
    ("PART", r"(?:mfr\.?\s*|manufacturer'?s?\s*)?part\s*(?:no\.?|number|#)", 3),
    ("PART", r"\bp/n\b", 3),
    ("PART", r"material\s*code", 2),
    ("ITEM_CODE", r"item\s*code", 3),
    ("TYPE", r"(?:valve\s*|actuator\s*|model\s*)?type\s*(?:no\.?)?\s*[:=]", 1),
    ("MANUFACTURER", r"(?:actuator\s*)?(?:manufacturer|mfg\.?\s*name\.?|mfr\.?|vendor\s*name)", 3),
    ("MANUFACTURER", r"(?:actuator\s*)?make\b", 1),
    ("DESCRIPTION", r"(?:product\s*|item\s*|equipment\s*)?description", 2),
]
LABELS = [(f, re.compile(r"(?<![a-z])" + rx + r"\s*[:.\-=]*\s*", re.I), s) for f, rx, s in LABELS]

CATALOG_CODE_RX = re.compile(r"catalog(?:ue)?\s*&\s*code\s*[:\-]?\s*(\S+)\s+([A-Za-z0-9_\-/.]{6,})", re.I)
SIZE_CLASS_RX = re.compile(
    r"(\d+(?:[./]\d+)?\s*(?:\"|''|in(?:ch)?)?\s*(?:class|cl\.?|#)\s*\d{2,4}\s+[A-Za-z ]{0,25}?valve)", re.I)


@dataclass
class Finding:
    file: str
    page: int
    method: str          # TEXT / OCR / OCR-TABLE / TEXT-TABLE
    field: str           # TAG / SERIAL / MODEL / PART / ITEM_CODE / MANUFACTURER / DESCRIPTION
    value: str
    raw: str = ""
    component: str = ""  # '' = not decided (page default applies)
    tag: str = ""        # tag the value was linked to
    link: str = ""       # row / line / page / serial-match / document
    strength: int = 2
    conf: float = 0.0    # OCR confidence 0-100 (100 for text)
    context: str = ""
    note: str = ""


def cut_value(rest: str) -> str:
    m = re.search(_STOP_WORDS, " " + rest, re.I)
    if m and m.start() > 0:
        rest = (" " + rest)[: m.start()]
    rest = rest.split("|")[0]
    return strip_junk(rest)


def extract_from_line(line, next_lines, conf, base):
    """Key:value extraction from one text line. base = dict(file, page, method)."""
    out = []
    raw_line = clean_text(line)
    low = raw_line.lower()
    if not raw_line:
        return out
    negative = any(n in low for n in NEGATIVE_CONTEXT)
    comp = component_of(low) or ""
    tags_here = sorted({t for t, _ in find_tags(raw_line)})
    line_tag = tags_here[0] if len(tags_here) == 1 else ""

    def add(field_, value, strength, rawv, note=""):
        if not value:
            return
        out.append(Finding(**base, field=field_, value=value, raw=rawv, component=comp,
                           tag=line_tag, link="line" if line_tag else "", strength=strength,
                           conf=conf, context=raw_line[:160], note=note))

    # Catalogue & Code  113  GA011BMAA12HSS_002
    m = CATALOG_CODE_RX.search(raw_line)
    if m:
        add("MODEL", strip_junk(m.group(1)), 1, m.group(0), "catalogue no.")
        add("PART", strip_junk(m.group(2)), 3, m.group(0), "catalogue code")

    # size / class description, e.g. 1" Class 150 Gate Valve
    m = SIZE_CLASS_RX.search(raw_line)
    if m and "type" not in low:
        add("DESCRIPTION", re.sub(r"\s+", " ", m.group(1)).strip(), 2, m.group(1))

    used_spans = []
    for fld, rx, strength in LABELS:
        for m in rx.finditer(raw_line):
            if any(a <= m.start() < b for a, b in used_spans):
                continue
            label = m.group(0).lower()
            rest = raw_line[m.end():]
            if fld == "SERIAL":
                if negative or re.search(r"\bpo\b|\bheat\b|\bgauge", low[:m.start() + 3]):
                    continue
                val = first_code_token(cut_value(rest) or rest[:40])
                if not val:
                    # label at end of line -> value on next line(s)
                    for nl in next_lines[:2]:
                        val = first_code_token(clean_text(nl))
                        if val:
                            break
                if val and val.lower() not in ("no", "nos"):
                    s = strength
                    if line_tag or re.search(r"valve|act|equipment|unit", label):
                        s = 3
                    used_spans.append((m.start(), m.end()))
                    add("SERIAL", val.upper(), s, m.group(0) + rest[:40])
            elif fld == "MANUFACTURER":
                if "customer" in low or "end user" in low:
                    continue
                if re.match(r"\W*s?\W*marking", rest, re.I) or "marking" in label:
                    continue
                v = cut_value(rest)
                v = re.split(r",|\bkancheepuram works\b|\bworks\b", v, flags=re.I)[0].strip()
                known = canon_manufacturer(v, fuzzy=True) or canon_manufacturer(rest[:60])
                if strength == 1 and not known:      # plain "Make" -> only known brands
                    continue
                if not known and (len(v) < 3 or not re.search(r"[A-Za-z]{3}", v)):
                    continue
                s = 3 if ("actuator" in label or strength >= 3) else strength
                used_spans.append((m.start(), m.end()))
                add("MANUFACTURER", known or v.upper(), s, m.group(0) + rest[:40])
            elif fld == "MODEL":
                if negative or re.search(r"make\s*/\s*model", low):
                    continue
                toks = []
                for t in cut_value(rest).split():
                    if toks and (len(strip_junk(t)) <= 1 or not re.search(r"[A-Za-z0-9]", t)):
                        break
                    toks.append(t)
                v = strip_junk(" ".join(toks[:4]))
                if len(v) < 2 or not re.search(r"\d", v):
                    continue
                used_spans.append((m.start(), m.end()))
                add("MODEL", v, strength, m.group(0) + rest[:40])
            elif fld in ("PART", "ITEM_CODE"):
                if negative and fld == "PART":
                    continue
                v = first_code_token(cut_value(rest) or rest[:40], min_digits=2, min_len=4)
                if not v:
                    for nl in next_lines[:1]:
                        if not strip_junk(cut_value(rest)):
                            v = first_code_token(clean_text(nl), min_digits=2, min_len=4)
                if v:
                    used_spans.append((m.start(), m.end()))
                    add(fld, v.upper(), strength, m.group(0) + rest[:40])
            elif fld == "TYPE":
                v = cut_value(rest)
                if len(v) < 2 or re.search(r"globe|check\s*valves|penetrant|removable|aqueous", v, re.I):
                    continue
                used_spans.append((m.start(), m.end()))
                add("TYPE", re.sub(r"\s+", " ", v)[:40], 1, m.group(0) + rest[:40])
            elif fld == "DESCRIPTION":
                v = cut_value(rest)
                if len(v) >= 4 and re.search(r"[A-Za-z]{3}", v):
                    used_spans.append((m.start(), m.end()))
                    add("DESCRIPTION", re.sub(r"\s+", " ", v)[:80], strength, m.group(0) + rest[:40])
    return out


# -----------------------------------------------------------------------------
#  OCR helpers
# -----------------------------------------------------------------------------
MAX_PAGE_PIXELS = 18_000_000     # ~A4 at 360 dpi; bigger pages (A3/A1 drawings) are rendered smaller
MAX_PAGE_SIDE = 6000
TESSERACT_TIMEOUT = 180          # seconds per OCR call; a stuck page is skipped instead of freezing everything


def to_gray(page, dpi):
    w_in, h_in = page.rect.width / 72.0, page.rect.height / 72.0
    if w_in > 0 and h_in > 0:
        dpi = min(dpi, (MAX_PAGE_PIXELS / (w_in * h_in)) ** 0.5, MAX_PAGE_SIDE / max(w_in, h_in))
        dpi = int(max(dpi, 72))
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
    return np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w).copy()


def rotate(img, deg):
    if deg == 90:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    if deg == 180:
        return cv2.rotate(img, cv2.ROTATE_180)
    if deg == 270:
        return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return img


def mean_conf(img, psm=6):
    d = pytesseract.image_to_data(img, config=f"--psm {psm}", output_type=pytesseract.Output.DICT,
                                  timeout=TESSERACT_TIMEOUT)
    c = [float(x) for x, t in zip(d["conf"], d["text"]) if float(x) >= 0 and len(t.strip()) > 1]
    return (sum(c) / len(c) if c else 0.0), len(c)


def best_rotation(img):
    """OSD guess, verified by OCR confidence (OSD alone is often wrong on forms)."""
    small = cv2.resize(img, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
    cands = [0]
    try:
        osd = pytesseract.image_to_osd(small, output_type=pytesseract.Output.DICT, timeout=TESSERACT_TIMEOUT)
        if osd.get("rotate", 0):
            cands.append(int(osd["rotate"]))
    except Exception:
        cands += [90, 270]
    if len(cands) == 1:
        return 0
    h, w = small.shape
    crop = lambda im: im[im.shape[0] // 6: im.shape[0] * 5 // 6, im.shape[1] // 6: im.shape[1] * 5 // 6]
    scores = {}
    for r in cands:
        c, n = mean_conf(crop(rotate(small, r)))
        scores[r] = c * min(1.0, n / 30.0)
    return max(scores, key=scores.get)


def ocr_lines(img, psm=6):
    """OCR -> (list of (line_text, mean_conf), list of word boxes (x, y, w, h, text, conf))."""
    d = pytesseract.image_to_data(img, config=f"--psm {psm}", output_type=pytesseract.Output.DICT,
                                  timeout=TESSERACT_TIMEOUT)
    lines = defaultdict(list)
    words = []
    for i, t in enumerate(d["text"]):
        if not t.strip():
            continue
        key = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        lines[key].append((d["left"][i], t, float(d["conf"][i])))
        words.append((d["left"][i], d["top"][i], d["width"][i], d["height"][i], t, float(d["conf"][i])))
    out = []
    for key in sorted(lines):
        ws = sorted(lines[key])
        txt = " ".join(t for _, t, _ in ws)
        cs = [c for _, _, c in ws if c >= 0]
        out.append((txt, sum(cs) / len(cs) if cs else 0.0))
    return out, words


def ocr_cell(crop, psm=7, scale=1.0, otsu=False):
    if scale != 1.0:
        crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    if otsu:
        crop = cv2.threshold(crop, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    d = pytesseract.image_to_data(crop, config=f"--psm {psm}", output_type=pytesseract.Output.DICT,
                                  timeout=60)
    words = [(t, float(c)) for t, c in zip(d["text"], d["conf"]) if t.strip()]
    txt = " ".join(t for t, _ in words)
    cs = [c for _, c in words if c >= 0]
    return txt.strip(), (sum(cs) / len(cs) if cs else 0.0)


# -----------------------------------------------------------------------------
#  table handling (grid detection on scanned pages)
# -----------------------------------------------------------------------------
def has_table_grid(img):
    """Quick test: does the page contain a ruled table with several rows?"""
    bw = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15)
    H, W = bw.shape
    hl = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, W // 8), 1)))
    rows = cv2.reduce(hl, 1, cv2.REDUCE_MAX).ravel() > 0
    n = int(np.sum(rows[1:] & ~rows[:-1]))
    return n >= 4


def detect_grid_cells(img):
    bw = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15)
    H, W = bw.shape
    hl = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, W // 30), 1)))
    vl = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, H // 60))))
    lines = cv2.dilate(hl | vl, np.ones((3, 3), np.uint8))
    cnts, _ = cv2.findContours(255 - lines, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(c) for c in cnts]
    boxes = [b for b in boxes if 22 < b[3] < H * 0.25 and 22 < b[2] < W * 0.7]

    def inside(i, o):
        return i[0] >= o[0] - 2 and i[1] >= o[1] - 2 and i[0] + i[2] <= o[0] + o[2] + 2 and i[1] + i[3] <= o[1] + o[3] + 2
    # drop outer frames that contain other cells (they are a whole table, not a cell)
    boxes = [o for o in boxes if sum(1 for i in boxes if i is not o and inside(i, o)) < 2]
    clean = img.copy()
    clean[lines > 0] = 255
    # rows
    boxes.sort(key=lambda b: (b[1] + b[3] / 2, b[0]))
    rows = []
    for b in boxes:
        cy = b[1] + b[3] / 2
        for r in rows:
            if abs(r["cy"] - cy) < min(b[3], r["h"]) * 0.45:
                r["cells"].append(b)
                break
        else:
            rows.append({"cy": cy, "h": b[3], "cells": [b]})
    for r in rows:
        r["cells"].sort()
    return rows, clean


def crop_cell(clean, b, pad=3):
    x, y, w, h = b
    c = clean[y + pad: y + h - pad, x + pad: x + w - pad]
    if c.size == 0:
        c = clean[y: y + h, x: x + w]
    return cv2.copyMakeBorder(c, 18, 18, 18, 18, cv2.BORDER_CONSTANT, value=255)


def _fuzzy_has(words, target, ratio=0.78):
    return any(difflib.SequenceMatcher(None, w, target).ratio() >= ratio for w in words)


def header_role(text):
    low = text.lower()
    words = re.findall(r"[a-z]+", low)
    if not words:
        return None
    if _fuzzy_has(words, "serial") or "s/n" in low or \
            (_fuzzy_has(words, "valve", 0.7) and any(w in ("sl", "si", "nos", "no", "s") for w in words)):
        return "SERIAL"
    if "tag" in words or _fuzzy_has(words, "tag", 0.9):
        return "TAG"
    if _fuzzy_has(words, "item") and _fuzzy_has(words, "code"):
        return "ITEM_CODE"
    if (_fuzzy_has(words, "part") and any(w in ("no", "number", "nos") for w in words)) or \
            (_fuzzy_has(words, "material") and _fuzzy_has(words, "code")):
        return "PART"
    if _fuzzy_has(words, "model") or _fuzzy_has(words, "catalogue") or _fuzzy_has(words, "figure"):
        return "MODEL"
    if _fuzzy_has(words, "description"):
        return "DESCRIPTION"
    if _fuzzy_has(words, "manufacturer") or "make" in words or _fuzzy_has(words, "vendor"):
        return "MANUFACTURER"
    return None


def vote(values):
    """values: list of (value, weight) -> best value, total weight of best, all."""
    agg = defaultdict(float)
    for v, w in values:
        if v:
            agg[v] += w
    if not agg:
        return None, 0.0, {}
    best = max(agg, key=agg.get)
    return best, agg[best], dict(agg)


def parse_cell_value(role, text):
    t = clean_text(text)
    if role == "TAG":
        tags = find_tags(t)
        return tags[0][0] if tags else None
    if role == "SERIAL":
        compact = re.sub(r"\s+", "", t)
        cands = [strip_junk(x) for x in re.findall(r"[A-Za-z0-9\-/]+", compact)]
        cands = [c for c in cands if code_like(c)]
        return max(cands, key=len).upper() if cands else None
    if role in ("PART", "ITEM_CODE"):
        compact = re.sub(r"\s+", "", t)
        cands = [strip_junk(x) for x in re.findall(r"[A-Za-z0-9\-/_.]+", compact)]
        cands = [c for c in cands if code_like(c, min_digits=1, min_len=4)]
        return max(cands, key=len).upper() if cands else None
    if role == "MANUFACTURER":
        return canon_manufacturer(t) or (strip_junk(t).upper() if len(strip_junk(t)) > 2 else None)
    v = re.sub(r"\s+", " ", strip_junk(t))
    return v if len(v) >= 3 else None


CELL_VARIANTS = [dict(psm=7), dict(psm=7, scale=2.0), dict(psm=6, scale=2.0), dict(psm=7, scale=2.0, otsu=True)]


def read_cell_voted(clean, box, role, mode="normal", seed_text=None, seed_conf=0.0):
    """Read one table cell several ways and let the readings vote.
    seed_text = what the full-page OCR already read inside this cell (free extra vote)."""
    crop = crop_cell(clean, box)
    vals, raws = [], []
    if seed_text:
        raws.append(seed_text)
        pv = parse_cell_value(role, seed_text)
        if pv:   # full-page OCR sees the cell in context -> counts double
            vals.append((pv, 2 * max(seed_conf, 20.0) / 100.0))
    if role in ("TAG", "SERIAL"):
        variants = CELL_VARIANTS if mode == "thorough" else CELL_VARIANTS[:3]
    else:
        variants = CELL_VARIANTS[:2] if mode == "thorough" else CELL_VARIANTS[:1]
    if mode == "fast":
        variants = variants[:1]
    for v in variants:
        try:
            txt, cf = ocr_cell(crop, **v)
        except Exception:
            continue
        raws.append(txt)
        pv = parse_cell_value(role, txt)
        if pv:
            vals.append((pv, max(cf, 20.0) / 100.0))
    if role == "TAG" and len(vals) >= 2:
        vals = [(segment_vote(vals), 1.0)] + vals
    best, w, allv = vote(vals)
    n = len(vals)
    agree = (w / sum(allv.values())) if allv else 0
    conf = 100.0 * agree * min(1.0, n / 2.0)
    return best, conf, " / ".join(raws[:3])


def segment_vote(vals):
    """Vote each part of the tag separately: 775-130 / 175-730 / 775-730 -> 775-730."""
    by_n = defaultdict(list)
    for v, w in vals:
        by_n[len(v.split("-"))].append((v.split("-"), w))
    n, group = max(by_n.items(), key=lambda kv: sum(w for _, w in kv[1]))
    parts = []
    for i in range(n):
        agg = defaultdict(float)
        for segs, w in group:
            agg[segs[i]] += w
        parts.append(max(agg, key=agg.get))
    return "-".join(parts)


def words_in_box(words, b):
    x, y, w, h = b
    sel = [(wx, wt, wc) for wx, wy, ww, wh, wt, wc in words
           if x <= wx + ww / 2 <= x + w and y <= wy + wh / 2 <= y + h]
    sel.sort()
    cs = [c for _, _, c in sel if c >= 0]
    return " ".join(t for _, t, _ in sel), (sum(cs) / len(cs) if cs else 0.0)


def extract_tables_ocr(img, words, base, mode="normal"):
    rows, clean = detect_grid_cells(img)
    if len(rows) < 2:
        return []
    # quick text of every cell = words the full-page OCR already found inside it
    quick = []
    for r in rows:
        cells = []
        for b in r["cells"]:
            txt, cf = words_in_box(words, b)
            cells.append({"box": b, "text": txt, "conf": cf})
        quick.append(cells)
    tag_rows = [i for i, cells in enumerate(quick) if any(find_tags(c["text"]) for c in cells)]
    if not tag_rows:
        return []
    # tag column = x position / width of the cells where a tag was read
    tag_cells = [c for i in tag_rows for c in quick[i] if find_tags(c["text"])]
    tag_xc = float(np.median([c["box"][0] + c["box"][2] / 2 for c in tag_cells]))
    tag_w = float(np.median([c["box"][2] for c in tag_cells]))

    # every row with a cell of similar size in that column is a candidate (the page
    # OCR often misses some rows - re-read them cell by cell)
    data_rows = []
    for i, cells in enumerate(quick):
        tcell = [c for c in cells if c["box"][0] <= tag_xc <= c["box"][0] + c["box"][2]
                 and 0.6 * tag_w <= c["box"][2] <= 1.6 * tag_w]
        if not tcell:
            continue
        tag, tconf, traw = read_cell_voted(clean, tcell[0]["box"], "TAG", mode, tcell[0]["text"], tcell[0]["conf"])
        if tag:
            data_rows.append((i, tcell[0], tag, tconf, traw))
    if not data_rows:
        return []
    first = data_rows[0][0]
    # header = up to 3 rows above first data row
    header_cells = []
    for hi in range(max(0, first - 3), first):
        for c in quick[hi]:
            txt, _ = ocr_cell(crop_cell(clean, c["box"]), psm=6, scale=2.0)
            role = header_role(txt) or header_role(c["text"])
            if role:
                header_cells.append((c["box"], role, txt))

    def role_of(box):
        x0, x1 = box[0], box[0] + box[2]
        best, ov = None, 0
        for hb, role, _ in header_cells:
            o = min(x1, hb[0] + hb[2]) - max(x0, hb[0])
            if o > ov and o > 0.4 * min(box[2], hb[2]):
                best, ov = role, o
        return best

    # content-based serial column if header did not reveal one
    has_serial_hdr = any(r == "SERIAL" for _, r, _ in header_cells)
    serial_col_x = None
    if not has_serial_hdr:
        colvals = defaultdict(list)
        for i, tcell, *_ in data_rows:
            for c in quick[i]:
                if c is tcell or role_of(c["box"]) in ("ITEM_CODE", "PART", "DESCRIPTION", "MODEL", "TAG"):
                    continue
                v = parse_cell_value("SERIAL", c["text"])
                colvals[round((c["box"][0] + c["box"][2] / 2) / 40)].append(v)
        for key, vs in colvals.items():
            good = [v for v in vs if v and re.search(r"[A-Za-z]", v) and re.search(r"\d{4}", v)]
            if len(vs) >= 2 and len(good) >= 0.6 * len(vs) and len(set(good)) >= 0.8 * len(good):
                serial_col_x = key
                break

    findings = []
    for i, tcell, tag, tconf, traw in data_rows:
        findings.append(Finding(**base, field="TAG", value=tag, raw=traw, tag=tag, link="row",
                                strength=3, conf=tconf, context="table row"))
        for c in quick[i]:
            if c is tcell:
                continue
            role = role_of(c["box"])
            note = ""
            if role is None and serial_col_x is not None and \
                    round((c["box"][0] + c["box"][2] / 2) / 40) == serial_col_x:
                role, note = "SERIAL", "serial column found by content"
            if role in (None, "TAG"):
                continue
            val, cf, raw = read_cell_voted(clean, c["box"], role, mode, c["text"], c["conf"])
            if val:
                findings.append(Finding(**base, field=role, value=val, raw=raw, tag=tag, link="row",
                                        strength=3 if not note else 2, conf=cf,
                                        context=f"table column '{role}'", note=note))
    return findings


def extract_tables_text(page, base):
    """Native-text PDFs: PyMuPDF table finder."""
    out = []
    try:
        tabs = page.find_tables()
    except Exception:
        return out
    for tab in tabs:
        try:
            data = tab.extract()
        except Exception:
            continue
        if not data or len(data) < 2:
            continue
        header_idx, roles = None, {}
        for hi, row in enumerate(data[:4]):
            r = {j: header_role(str(c or "")) for j, c in enumerate(row)}
            if "TAG" in r.values():
                header_idx, roles = hi, {j: v for j, v in r.items() if v}
                break
        if header_idx is None:
            continue
        tag_col = [j for j, v in roles.items() if v == "TAG"][0]
        for row in data[header_idx + 1:]:
            tags = find_tags(str(row[tag_col] or ""))
            if not tags:
                continue
            tag = tags[0][0]
            out.append(Finding(**base, field="TAG", value=tag, raw=str(row[tag_col]), tag=tag,
                               link="row", strength=3, conf=100, context="table row"))
            for j, role in roles.items():
                if role == "TAG" or j >= len(row):
                    continue
                v = parse_cell_value(role, str(row[j] or ""))
                if v:
                    out.append(Finding(**base, field=role, value=v, raw=str(row[j]), tag=tag, link="row",
                                       strength=3, conf=100, context=f"table column '{role}'"))
    return out


# -----------------------------------------------------------------------------
#  page processing (runs in worker processes)
# -----------------------------------------------------------------------------
_PDF_LOCK = threading.Lock()
_STANDALONE = False          # set to True by main()


def process_page(path, pno, dpi=300, mode="normal", tesseract_cmd=None):
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    fname = os.path.basename(path)
    res = {"file": fname, "path": path, "page": pno + 1, "method": "", "rotation": 0,
           "lines": [], "findings": [], "error": ""}
    try:
        lines, table_f, img = [], [], None
        with _PDF_LOCK:                      # PDF library is not thread-safe -> one page at a time
            doc = pymupdf.open(path)
            page = doc[pno]
            native = page.get_text("text") or ""
            alnum = sum(ch.isalnum() for ch in native)
            if alnum >= 80:
                res["method"] = "TEXT"
                blocks = page.get_text("dict")["blocks"]
                spans = []
                for b in blocks:
                    for l in b.get("lines", []):
                        t = " ".join(s["text"] for s in l["spans"]).strip()
                        if t:
                            spans.append((round(l["bbox"][1] / 4), l["bbox"][0], t))
                rowmap = defaultdict(list)
                for y, x, t in spans:
                    rowmap[y].append((x, t))
                for y in sorted(rowmap):
                    lines.append((" ".join(t for _, t in sorted(rowmap[y])), 100.0))
                base = dict(file=fname, page=pno + 1, method="TEXT-TABLE")
                table_f = extract_tables_text(page, base)
            else:
                res["method"] = "OCR"
                img = to_gray(page, dpi)
            doc.close()
        if img is not None:
            rot = best_rotation(img)
            res["rotation"] = rot
            img = rotate(img, rot)
            lines, words = ocr_lines(img, psm=6)
            if mode == "thorough":
                lines2, _ = ocr_lines(img, psm=4)
                seen = {l for l, _ in lines}
                lines += [(l, c) for l, c in lines2 if l not in seen]
            # cell-by-cell table reading only where several tags are listed
            n_tags = len({t for l, _ in lines for t, _ in find_tags(l)})
            tag_header = any(re.search(r"\btag\b", l, re.I) and
                             re.search(r"\b(?:qty|item|serial|s[il1]\.?\s*no|description|code|gad)\b", l, re.I)
                             for l, _ in lines)
            if n_tags >= 1 and mode != "fast" and (n_tags >= 2 or tag_header or has_table_grid(img)):
                base = dict(file=fname, page=pno + 1, method="OCR-TABLE")
                table_f = extract_tables_ocr(img, words, base, mode)
        res["lines"] = lines
        base = dict(file=fname, page=pno + 1, method=res["method"])
        f = []
        for i, (l, c) in enumerate(lines):
            for tag, raw in find_tags(l):
                f.append(Finding(**base, field="TAG", value=tag, raw=raw, tag=tag, link="line",
                                 strength=2, conf=c, context=clean_text(l)[:160]))
            f += extract_from_line(l, [x for x, _ in lines[i + 1:i + 3]], c, base)
        res["findings"] = [asdict(x) for x in f + table_f]
    except MemoryError:
        res["error"] = "MemoryError: page too large for available memory - page skipped"
    except Exception as e:
        res["error"] = f"{type(e).__name__}: {e}"
        res["trace"] = traceback.format_exc()
    return res


# -----------------------------------------------------------------------------
#  batch-level post processing
# -----------------------------------------------------------------------------
def seg_snap_map(tag_weights):
    """Repair OCR digit confusions in tag segments using segment frequency."""
    segs = defaultdict(Counter)
    for t, w in tag_weights.items():
        parts = t.split("-")
        for i, p in enumerate(parts):
            segs[(len(parts), i)][p] += w
    mapping = {}
    for t in tag_weights:
        parts = t.split("-")
        new = []
        for i, p in enumerate(parts):
            cnt = segs[(len(parts), i)]
            best = p
            for q, wq in cnt.most_common():
                if q == p or len(q) != len(p):
                    continue
                diff = [(a, b) for a, b in zip(p, q) if a != b]
                if diff and len(diff) <= 1 and all(d in _CONFUSABLE for d in diff) and wq >= 2.5 * cnt[p] and wq >= 1.5:
                    best = q
                    break
            new.append(best)
        nt = "-".join(new)
        if nt != t:
            mapping[t] = nt
    return mapping


def levenshtein(a, b):
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def build_tag_map(all_tags_weighted, master_tags=None, tag_pages=None):
    """raw tag -> final tag (snap to master list, else segment / fuzzy consensus)."""
    mapping = {}
    if master_tags:
        # 1) usual OCR consensus clean-up first, 2) then snap to the tag list - but only when the
        #    difference is a typical OCR mix-up (7/1, 0/8 ...). Real, different tags are never merged.
        mapping = build_tag_map(all_tags_weighted, None, tag_pages)
        mset = set(master_tags)
        for t in all_tags_weighted:
            c = mapping.get(t, t)
            if t in mset or c in mset:
                if t in mset:
                    mapping.pop(t, None)
                continue
            cands = []
            for m in mset:
                if len(m) != len(c):
                    continue
                diff = [(x, y) for x, y in zip(c, m) if x != y]
                if 0 < len(diff) <= 2 and all(d in _CONFUSABLE for d in diff):
                    cands.append((len(diff), m))
            cands.sort()
            if cands and (len(cands) == 1 or cands[0][0] < cands[1][0]):
                mapping[t] = cands[0][1]
        return mapping
    # 1. segment snapping
    mapping.update(seg_snap_map(all_tags_weighted))
    merged = Counter()
    for t, w in all_tags_weighted.items():
        merged[mapping.get(t, t)] += w
    # 2. weak tag within distance 2 of a much stronger tag -> merge
    strong = [t for t, _ in merged.most_common()]
    tag_pages = tag_pages or {}
    pages_of = lambda x: len({pg for k, v in mapping.items() if v == x for pg in tag_pages.get(k, ())} |
                             set(tag_pages.get(x, ())))
    for t in sorted(merged, key=merged.get):
        for s in strong:
            if s == t:
                continue
            stronger = merged[s] >= 2 * merged[t] or \
                (merged[s] >= 1.2 * merged[t] and pages_of(t) <= 1 and pages_of(s) >= 2)
            if not stronger:
                continue
            if len(s) == len(t) and levenshtein(s, t) <= 2 and \
                    all((a, b) in _CONFUSABLE for a, b in zip(t, s) if a != b):
                for k, v in list(mapping.items()) + [(t, t)]:
                    if v == t:
                        mapping[k] = s
                mapping[t] = s
                break
    return mapping


def learn_serial_formats(findings):
    shapes = Counter()
    for f in findings:
        if f["field"] == "SERIAL" and f["strength"] >= 3 and f["tag"] and f["conf"] >= 50:
            shapes[shape_of(f["value"])] += 1
    learned = {s for s, n in shapes.items() if n >= 2 and re.search(r"9{3}", s)}
    learned |= set(SERIAL_FORMATS)
    # '999999999' next to '99C999999': the digit-only one is just the C misread as 0
    drop = set()
    for a in learned:
        for b in learned:
            if a != b and len(a) == len(b):
                diff = [(x, y) for x, y in zip(a, b) if x != y]
                if diff and all(x == "9" and y.isalpha() for x, y in diff):
                    drop.add(a)
    return learned - drop


def repair_to_format(value, formats):
    """Fix confusions so value fits a known/learned format. Returns (new, changed)."""
    v = re.sub(r"\s+", "", value.upper())
    if shape_of(v) in formats:
        return v, v != value
    for fmt in formats:
        if len(fmt) != len(v):
            continue
        out, ok, lit_changes = [], True, 0
        for c, f in zip(v, fmt):
            if f == "9":
                d = _DIGIT_FIX.get(c, c)
                if not d.isdigit():
                    ok = False
                    break
                out.append(d)
            elif f.isalpha():
                if c != f:
                    lit_changes += 1
                out.append(f)
            else:
                out.append(f)
        if ok and lit_changes <= 1:
            nv = "".join(out)
            return nv, nv != value
    # junk around the value (e.g. '4950123255' from a tick mark): look for a window
    for fmt in sorted(formats, key=len, reverse=True):
        L = len(fmt)
        if len(v) > L and len(v) <= L + 3:
            for start in range(len(v) - L, -1, -1):
                nv, ch = repair_to_format(v[start:start + L], {fmt})
                if shape_of(nv) == fmt:
                    return nv, True
    return v, False


def norm_cmp(v):
    return re.sub(r"[^A-Z0-9]", "", v.upper())


def postprocess(page_results, master_tags=None):
    findings = [dict(f) for r in page_results for f in r["findings"]]
    pages_text = {(r["file"], r["page"]): " ".join(l for l, _ in r["lines"]) for r in page_results}
    page_header = {}
    page_is_cert = {}
    skip_rx = re.compile(r"customer|end\s*user|mfg|manufacturer|make|maker|vendor|actuator|supplier|"
                         r"registered office|l&t house", re.I)
    for r in page_results:
        key = (r["file"], r["page"])
        brands = Counter()
        first_pos = {}
        for i, (l, _) in enumerate(r["lines"]):
            if skip_rx.search(l):
                continue
            b = canon_manufacturer(l, fuzzy=True)
            if b:
                brands[b] += 1
                first_pos.setdefault(b, i)
        page_header[key] = max(brands, key=lambda b: (brands[b], -first_pos[b])) if brands else None
        low = pages_text[key].lower()
        page_is_cert[key] = any(k in low for k in COMPONENT_CERT_KEYWORDS)

    # --- 1. tag consolidation --------------------------------------------------
    tw = Counter()
    for f in findings:
        if f["field"] == "TAG":
            tw[f["value"]] += (1.0 if f["link"] == "row" else 0.7) * max(0.3, f["conf"] / 100)
    tpages = defaultdict(set)
    for f in findings:
        if f["field"] == "TAG":
            tpages[f["value"]].add((f["file"], f["page"]))
    tag_map = build_tag_map(tw, master_tags, tpages)
    if master_tags:
        mset = set(master_tags)
        for f in findings:
            if f["field"] == "TAG" and tag_map.get(f["value"], f["value"]) not in mset:
                f["note"] = (f["note"] + "; not in tag list").strip("; ")
    for f in findings:
        for k in ("tag",):
            if f[k] and f[k] in tag_map:
                new = tag_map[f[k]]
                if new is None:
                    f["note"] = (f["note"] + "; tag not in master list").strip("; ")
                    f[k] = ""
                else:
                    f["note"] = (f["note"] + f"; tag OCR '{f[k]}' corrected").strip("; ")
                    f[k] = new
        if f["field"] == "TAG":
            f["value"] = f["tag"] or f["value"]

    # page tags
    page_tags = defaultdict(set)
    for f in findings:
        if f["field"] == "TAG" and f["tag"]:
            page_tags[(f["file"], f["page"])].add(f["tag"])

    # --- 2. page default component ---------------------------------------------
    def default_comp(key):
        hb = page_header.get(key)
        if hb in ACTUATOR_BRANDS:
            return "ACTUATOR"
        return DEFAULT_COMPONENT

    for f in findings:
        if not f["component"]:
            f["component"] = default_comp((f["file"], f["page"]))
        if f["field"] == "MANUFACTURER" and f["component"] == "ACTUATOR" and \
                page_header.get((f["file"], f["page"])) in ACTUATOR_BRANDS and \
                f["value"] != page_header.get((f["file"], f["page"])) and "actuator" not in f["context"].lower():
            f["component"] = DEFAULT_COMPONENT
        if f["field"] == "MANUFACTURER" and f["component"] == DEFAULT_COMPONENT and \
                f["value"] in ACTUATOR_BRANDS and f["strength"] < 3:
            f["component"] = "ACTUATOR"

    # page-header brand as a weak manufacturer finding
    for r in page_results:
        key = (r["file"], r["page"])
        hb = page_header.get(key)
        if hb and not page_is_cert[key]:
            findings.append(dict(asdict(Finding(file=r["file"], page=r["page"], method=r["method"],
                                                field="MANUFACTURER", value=hb, raw=hb,
                                                component=default_comp(key), strength=1, conf=80,
                                                context="brand in page header"))))

    # --- 3. link by page (exactly one tag on the page) -------------------------
    for f in findings:
        key = (f["file"], f["page"])
        if f["field"] != "TAG" and not f["tag"] and len(page_tags[key]) == 1 and not page_is_cert[key]:
            f["tag"] = next(iter(page_tags[key]))
            f["link"] = "page"

    # --- 4. serial formats + repair -------------------------------------------
    formats = learn_serial_formats(findings)
    for f in findings:
        if f["field"] == "SERIAL" and formats:
            nv, changed = repair_to_format(f["value"], formats)
            if changed and nv != f["value"]:
                f["note"] = (f["note"] + f"; OCR '{f['value']}' fixed to format").strip("; ")
                f["value"] = nv

    # strong serials per (tag, component)
    strong_serial = {}
    for f in findings:
        if f["field"] == "SERIAL" and f["tag"] and f["strength"] >= 3 and \
                (not formats or shape_of(f["value"]) in formats):
            strong_serial.setdefault(norm_cmp(f["value"]), (f["tag"], f["component"]))

    # --- 5. link pages with no tag but containing a known serial -------------
    page_serial_link = {}
    for key, txt in pages_text.items():
        if page_tags[key] or page_is_cert[key]:
            continue
        toks = set()
        for t in re.split(r"\s+", txt):
            t = strip_junk(t)
            if len(t) >= 5:
                toks.add(norm_cmp(repair_to_format(t, formats)[0]) if formats else norm_cmp(t))
        compact = norm_cmp(txt)
        hits = {v for v in strong_serial if v in toks or (len(v) >= 7 and v in compact)}
        linked = {strong_serial[v][0] for v in hits}
        if len(linked) == 1:
            page_serial_link[key] = (linked.pop(), hits)
    for f in findings:
        key = (f["file"], f["page"])
        if f["field"] != "TAG" and not f["tag"] and key in page_serial_link:
            f["tag"] = page_serial_link[key][0]
            f["link"] = "serial-match"
    meth = {(r["file"], r["page"]): r["method"] for r in page_results}
    for key, (tag, hits) in page_serial_link.items():
        for v in hits:
            t, comp = strong_serial[v]
            findings.append(asdict(Finding(file=key[0], page=key[1], method=meth[key], field="SERIAL",
                                           value=v, raw=v, component=comp, tag=t, link="serial-match",
                                           strength=2, conf=80, context="serial printed on this page")))

    # --- 6. document-level manufacturer for multi-tag documents ----------------
    by_file_tags = defaultdict(set)
    for f in findings:
        if f["tag"]:
            by_file_tags[f["file"]].add(f["tag"])
    doc_mfr = defaultdict(Counter)
    for f in findings:
        if f["field"] == "MANUFACTURER" and not f["tag"] and f["component"] == DEFAULT_COMPONENT \
                and not page_is_cert[(f["file"], f["page"])]:
            if f["value"] in KNOWN_MANUFACTURERS.values():
                doc_mfr[f["file"]][f["value"]] += f["strength"]
    extra = []
    for file, cnt in doc_mfr.items():
        val, top = cnt.most_common(1)[0]
        if top >= 0.8 * sum(cnt.values()):
            src = [f for f in findings if f["file"] == file and f["field"] == "MANUFACTURER"
                   and f["value"] == val and not f["tag"]]
            for t in by_file_tags[file]:
                for s in src:
                    e = dict(s)
                    e["tag"], e["link"] = t, "document"
                    extra.append(e)
    findings += extra

    # --- 7. drop serials that don't fit the learned formats when weakly linked -
    for f in findings:
        if f["field"] == "SERIAL" and formats and shape_of(f["value"]) not in formats \
                and f["link"] not in ("row", "line") and f["strength"] < 3:
            f["note"] = (f["note"] + "; ignored: not in learned serial format").strip("; ")
            f["ignored"] = True
    return findings, formats


LINK_W = {"row": 1.0, "line": 1.0, "page": 0.85, "serial-match": 0.7, "document": 0.4, "": 0.0}


def consolidate(findings, formats=()):
    groups = defaultdict(list)
    tags_seen = defaultdict(list)
    for f in findings:
        if not f.get("tag") or f.get("ignored"):
            continue
        if f["field"] == "TAG":
            tags_seen[f["tag"]].append(f)
            continue
        w = LINK_W.get(f["link"], 0.3) * (f["strength"] / 3.0) * max(0.3, f["conf"] / 100.0)
        if f["method"].startswith("TEXT"):
            w *= 1.2
        if f["field"] == "SERIAL" and formats:
            w *= 1.5 if shape_of(f["value"]) in formats else 0.4
        groups[(f["tag"], f["component"], f["field"])].append((f, w))

    rows = []
    comps_per_tag = defaultdict(set)
    for (tag, comp, _), _ in groups.items():
        comps_per_tag[tag].add(comp)
    for tag in tags_seen:
        if not comps_per_tag[tag]:
            comps_per_tag[tag].add(DEFAULT_COMPONENT)

    def pick(tag, comp, fld):
        """best value for one field + the findings that support it + review notes"""
        items = groups.get((tag, comp, fld), [])
        agg, srcs = defaultdict(float), defaultdict(list)
        for f, w in items:
            agg[f["value"].strip()] += w
            srcs[f["value"].strip()].append(f)
        if not agg:
            return "", [], []
        best = max(agg, key=agg.get)
        share = agg[best] / sum(agg.values())
        bf = srcs[best]
        pages = {(f["file"], f["page"]) for f in bf}
        notes = []
        others = [v for v in sorted(agg, key=agg.get, reverse=True) if v != best][:3]
        if share < 0.66 and others:
            notes.append(f"other readings: {', '.join(others)}")
        elif len(pages) == 1 and max(f["conf"] for f in bf) < 75 and not bf[0]["method"].startswith("TEXT"):
            notes.append("single low-confidence read")
        return best, bf, notes

    def pages_text(flist):
        """'doc.pdf: p1, p2 | other.pdf: p30'"""
        by_doc = defaultdict(set)
        for f in flist:
            by_doc[f["file"]].add(f["page"])
        return " | ".join(f"{doc}: " + ", ".join(f"p{p}" for p in sorted(pg)) for doc, pg in sorted(by_doc.items()))

    for tag in sorted(comps_per_tag):
        comps = sorted(comps_per_tag[tag], key=lambda c: (c != DEFAULT_COMPONENT, c))
        for comp in comps:
            row = {"Tag Number": tag, "Equipment": comp}
            review = []

            mfr, _, n = pick(tag, comp, "MANUFACTURER")
            row["Manufacturer"] = mfr
            review += [f"Manufacturer: {x}" for x in n]

            # Model: model number / model name first, then item code, then type
            for fld, label in (("MODEL", ""), ("ITEM_CODE", "from Item Code"), ("TYPE", "from Type")):
                val, _, n = pick(tag, comp, fld)
                if val:
                    row["Model Number"] = val
                    row["Model Taken From"] = label or "Model"
                    review += [f"Model: {x}" for x in n]
                    break
            else:
                row["Model Number"], row["Model Taken From"] = "", ""

            ser, ser_src, n = pick(tag, comp, "SERIAL")
            row["Serial Number"] = ser
            row["Serial Number - Source (Document / Page)"] = pages_text(ser_src)
            row["_serial_src"] = sorted({(f["file"], f["page"]) for f in ser_src})
            # other serials for the same tag that look real (right format, seen on 2+ pages)
            alt = defaultdict(set)
            for f, _ in groups.get((tag, comp, "SERIAL"), []):
                v = f["value"].strip()
                if v != ser:
                    alt[v].add((f["file"], f["page"]))
            row["_alt_serials"] = {v: sorted(pg) for v, pg in alt.items()
                                   if len(pg) >= 2 and (not formats or shape_of(v) in formats)}
            review += [f"Serial: {x}" for x in n]
            if ser and len({(f["file"], f["page"]) for f in ser_src}) == 1 and \
                    any("fixed to format" in f["note"] for f in ser_src):
                review.append("Serial: OCR auto-corrected")
            if not ser:
                review.append("Serial not found")

            part, _, n = pick(tag, comp, "PART")
            row["Part Number"] = part
            review += [f"Part No.: {x}" for x in n]

            tag_src = tags_seen.get(tag, [])
            row["Tag Number - Source (Document / Page)"] = pages_text(tag_src)
            if tag_src and all("not in tag list" in f.get("note", "") for f in tag_src):
                review.append("Tag not in your tag list - check it")
            if len({(f["file"], f["page"]) for f in tag_src}) <= 1 and \
                    any("corrected" in f.get("note", "") for f in tag_src):
                review.append("Tag: OCR auto-corrected")
            row["Review"] = " | ".join(review)
            rows.append(row)
    return rows


def find_duplicates(rows, findings):
    """Returns (rows without duplicates, list of duplicate records for the Duplicates sheet)."""
    dups = []

    def pages_by_doc(pairs):
        by = defaultdict(set)
        for fn, pg in pairs:
            by[fn].add(pg)
        return by

    def add(kind, value, tag, comp, pairs, detail=""):
        if not pairs:
            dups.append({"Duplicate Type": kind, "Value": value, "Tag Number": tag, "Equipment": comp,
                         "Document Name": "", "Page Number(s)": "", "Details": detail})
        for doc, pgs in sorted(pages_by_doc(pairs).items()):
            dups.append({"Duplicate Type": kind, "Value": value, "Tag Number": tag, "Equipment": comp,
                         "Document Name": doc, "Page Number(s)": ", ".join(str(p) for p in sorted(pgs)),
                         "Details": detail})

    def note(row, txt):
        row["Review"] = (row.get("Review") + " | " if row.get("Review") else "") + txt

    # 1. same Tag + Equipment twice in the main sheet -> keep one row
    unique, seen = [], {}
    for r in rows:
        key = (r["Tag Number"], r["Equipment"])
        if key in seen:
            first = seen[key]
            for k, v in r.items():
                if v and not first.get(k):
                    first[k] = v
            add("Tag repeated in result", r["Tag Number"], r["Tag Number"], r["Equipment"],
                r.get("_serial_src", []), "merged into one row")
            continue
        seen[key] = r
        unique.append(r)
    rows = unique

    # 2. same serial number on different tags / equipment
    by_serial = defaultdict(list)
    for r in rows:
        if r.get("Serial Number"):
            by_serial[r["Serial Number"]].append(r)
    for ser, rs in sorted(by_serial.items()):
        if len(rs) < 2:
            continue
        tags = [f"{r['Tag Number']} ({r['Equipment']})" for r in rs]
        for r in rs:
            others = [t for t in tags if not t.startswith(r["Tag Number"] + " (" + r["Equipment"])]
            add("Same serial number on more than one tag", ser, r["Tag Number"], r["Equipment"],
                r.get("_serial_src", []), "also on: " + ", ".join(others))
            note(r, f"DUPLICATE serial - also on {', '.join(others)}")

    # 3. one tag with two different serial numbers (both read clearly on several pages)
    for r in rows:
        for alt, pairs in sorted(r.get("_alt_serials", {}).items()):
            add("Same tag with different serial numbers", r["Serial Number"], r["Tag Number"], r["Equipment"],
                r.get("_serial_src", []), f"serial used in result")
            add("Same tag with different serial numbers", alt, r["Tag Number"], r["Equipment"],
                pairs, f"other serial found for this tag (result uses {r['Serial Number']})")
            note(r, f"CONFLICT: another serial {alt} also found for this tag")

    # 4. same tag listed more than once in one table on the same page
    row_count = Counter()
    for f in findings:
        if f["field"] == "TAG" and f["tag"] and f["link"] == "row":
            row_count[(f["tag"], f["file"], f["page"])] += 1
    for (tag, fn, pg), n in sorted(row_count.items()):
        if n > 1:
            add("Tag listed more than once on the same page", tag, tag, "", [(fn, pg)],
                f"{n} table rows show this tag")
            for r in rows:
                if r["Tag Number"] == tag:
                    note(r, f"Tag appears in {n} rows on {fn} p{pg}")
    return rows, dups


# -----------------------------------------------------------------------------
#  Excel output
# -----------------------------------------------------------------------------
MAIN_COLUMNS = ["Tag Number", "Equipment", "Manufacturer", "Model Number", "Model Taken From",
                "Serial Number", "Part Number",
                "Tag Number - Source (Document / Page)", "Serial Number - Source (Document / Page)", "Review"]


def write_excel(rows, findings, file_log, out_path, formats, dups=()):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(bold=True, color="FFFFFF")
    review_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def sheet(ws, headers, data, widths):
        ws.append(headers)
        for c in ws[1]:
            c.fill, c.font = hdr_fill, hdr_font
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.border = border
        for d in data:
            ws.append([d.get(h, "") for h in headers])
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.border = border
                c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions

    # 1. main sheet
    ws = wb.active
    ws.title = "SAP_Equipment"
    sheet(ws, MAIN_COLUMNS, rows, [20, 12, 20, 26, 15, 16, 22, 45, 45, 45])
    rev_col = MAIN_COLUMNS.index("Review")
    for r in ws.iter_rows(min_row=2):
        if r[rev_col].value:
            for c in r:
                c.fill = review_fill
        r[0].font = Font(bold=True)

    # 2. where every tag was found: Tag / Document / Page
    ws2 = wb.create_sheet("All_Findings")
    seen = sorted({(f["tag"], f["file"], f["page"]) for f in findings if f["field"] == "TAG" and f["tag"]})
    data = [{"Tag Number": t, "Document Name": fn, "Page Number": pg} for t, fn, pg in seen]
    sheet(ws2, ["Tag Number", "Document Name", "Page Number"], data, [22, 60, 13])

    # 3. duplicates
    ws_d = wb.create_sheet("Duplicates", 1)
    dh = ["Duplicate Type", "Value", "Tag Number", "Equipment", "Document Name", "Page Number(s)", "Details"]
    if dups:
        sheet(ws_d, dh, list(dups), [36, 20, 20, 12, 55, 14, 45])
        red = PatternFill("solid", fgColor="F8CBAD")
        for r in ws_d.iter_rows(min_row=2):
            r[0].fill = red
    else:
        sheet(ws_d, dh, [{"Duplicate Type": "No duplicates found"}], [36, 20, 20, 12, 55, 14, 45])

    # 4. files processed / errors
    ws3 = wb.create_sheet("Files_Log")
    sheet(ws3, ["file", "pages", "tags_found", "errors"],
          [{k: d[k] for k in ("file", "pages", "tags_found", "errors")} for d in file_log], [60, 8, 70, 50])
    wb.save(out_path)


# -----------------------------------------------------------------------------
#  main
# -----------------------------------------------------------------------------
def find_tesseract(user_path=None):
    if user_path:
        return user_path
    from shutil import which
    if which("tesseract"):
        return None
    for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
              os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
              "/usr/local/bin/tesseract", "/opt/homebrew/bin/tesseract"]:
        if os.path.exists(p):
            return p
    return None


def load_taglist(path):
    tags = []
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            col = 0
            for j, h in enumerate(rows[0]):
                if h and "tag" in str(h).lower():
                    col = j
                    break
            for r in rows:
                if r and col < len(r) and r[col]:
                    tags += [t for t, _ in find_tags(str(r[col]))] or []
    else:
        with open(path, encoding="utf-8", errors="ignore") as fh:
            for line in fh:
                tags += [t for t, _ in find_tags(line)]
    return sorted(set(tags))


def collect_pdfs(inputs, recursive=False):
    out = []
    for p in inputs:
        if os.path.isdir(p):
            pat = os.path.join(p, "**", "*.pdf") if recursive else os.path.join(p, "*.pdf")
            out += glob.glob(pat, recursive=recursive)
            pat2 = pat[:-3] + "PDF"
            out += [x for x in glob.glob(pat2, recursive=recursive) if x not in out]
        elif os.path.isfile(p) and p.lower().endswith(".pdf"):
            out.append(p)
        else:
            out += [x for x in glob.glob(p) if x.lower().endswith(".pdf")]
    seen, res = set(), []
    for x in out:
        a = os.path.abspath(x)
        if a not in seen:
            seen.add(a)
            res.append(a)
    return res


def run(inputs, output=None, taglist=None, recursive=False, mode="normal", dpi=300, workers=None,
        tesseract=None, log=print, progress=None, use_processes=None):
    tcmd = find_tesseract(tesseract)
    if tcmd:
        pytesseract.pytesseract.tesseract_cmd = tcmd
    pdfs = collect_pdfs(inputs, recursive)
    if not pdfs:
        log("No PDF files found.")
        return None
    master = load_taglist(taglist) if taglist else None
    if master:
        log(f"Master tag list: {len(master)} tags")
    jobs = []
    for p in pdfs:
        try:
            n = len(pymupdf.open(p))
        except Exception as e:
            log(f"!! cannot open {p}: {e}")
            continue
        jobs += [(p, i) for i in range(n)]
    log(f"{len(pdfs)} PDF file(s), {len(jobs)} page(s) to read ...")
    if use_processes is None:
        use_processes = _STANDALONE      # True when run directly, False when used by a web page / notebook
    workers = workers or max(1, min(4, (os.cpu_count() or 2) - (0 if IN_COLAB else 1)))
    results = {}

    def collect(r):
        results[(r["path"], r["page"] - 1)] = r
        if r["error"]:
            log(f"  !! {r['file']} p{r['page']}: {r['error']}")
        if progress:
            progress(len(results), len(jobs))
        if len(results) % 5 == 0 or len(results) == len(jobs):
            log(f"  {len(results)}/{len(jobs)} pages done")

    # A) separate processes (fastest). Some setups kill or block them - web apps, notebooks,
    #    antivirus, low memory. Then we simply continue with B) threads in this process.
    if use_processes and workers > 1 and len(jobs) > 1:
        try:
            ctx = None
            if sys.platform.startswith("linux"):
                import multiprocessing
                ctx = multiprocessing.get_context("fork")
            with ProcessPoolExecutor(max_workers=workers, mp_context=ctx) as ex:
                futs = {ex.submit(process_page, p, i, dpi, mode, tcmd): (p, i) for p, i in jobs}
                for fu in as_completed(futs):
                    collect(fu.result())
        except Exception as e:
            log(f"  (parallel processes stopped: {type(e).__name__} - continuing in safe mode, "
                f"{len(jobs) - len(results)} page(s) left)")

    # B) safe mode: threads inside this process (OCR still runs in parallel)
    todo = [(p, i) for p, i in jobs if (p, i) not in results]
    if todo:
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(process_page, p, i, dpi, mode, tcmd): (p, i) for p, i in todo}
            for fu in as_completed(futs):
                p, i = futs[fu]
                try:
                    collect(fu.result())
                except Exception as e:
                    collect({"file": os.path.basename(p), "path": p, "page": i + 1, "method": "",
                             "rotation": 0, "lines": [], "findings": [], "error": f"{type(e).__name__}: {e}"})
    results = list(results.values())
    failed = [r for r in results if r["error"]]
    if failed:
        log(f"  {len(failed)} page(s) could not be read (see Files_Log sheet): "
            + ", ".join(f"{r['file']} p{r['page']}" for r in failed[:10]))
    results.sort(key=lambda r: (r["file"], r["page"]))

    findings, formats = postprocess(results, master)
    rows = consolidate(findings, formats)
    rows, dups = find_duplicates(rows, findings)

    file_log = []
    for p in pdfs:
        fn = os.path.basename(p)
        rs = [r for r in results if r["file"] == fn]
        tags = sorted({f["tag"] for f in findings if f["file"] == fn and f["field"] == "TAG" and f["tag"]})
        file_log.append({
            "file": fn, "pages": len(rs),
            "text_pages": sum(r["method"] == "TEXT" for r in rs),  # noqa
            "ocr_pages": sum(r["method"] == "OCR" for r in rs),
            "rotated_pages": ", ".join(f"p{r['page']}:{r['rotation']}°" for r in rs if r["rotation"]),
            "tags_found": ", ".join(tags),
            "errors": "; ".join(f"p{r['page']}: {r['error']}" for r in rs if r["error"]),
        })
    if not output:
        output = os.path.join(os.path.dirname(pdfs[0]),
                              f"SAP_Equipment_Extract_{datetime.datetime.now():%Y%m%d_%H%M}.xlsx")
    write_excel(rows, findings, file_log, output, formats, dups)
    if dups:
        log(f"{len({(d['Duplicate Type'], d['Value']) for d in dups})} duplicate issue(s) - see the 'Duplicates' sheet.")
    else:
        log("No duplicates found.")
    log(f"\n{len(rows)} equipment row(s) for {len({r['Tag Number'] for r in rows})} tag(s)"
        f"  ->  {output}")
    n_rev = sum(1 for r in rows if r["Review"])
    if n_rev:
        log(f"{n_rev} row(s) flagged for review (yellow).")
    return output


# -----------------------------------------------------------------------------
#  Upload window
# -----------------------------------------------------------------------------
def open_file(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def tesseract_ready(user_path=None):
    tcmd = find_tesseract(user_path)
    if tcmd:
        pytesseract.pytesseract.tesseract_cmd = tcmd
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


class UploadWindow:
    """Window where the user uploads (selects) PDFs every time the program starts."""

    def __init__(self, preset_files=(), preset_taglist=None, mode="normal"):
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        import threading
        import queue
        self.tk, self.ttk, self.fd, self.mb = tk, ttk, filedialog, messagebox
        self.threading, self.queue = threading, queue.Queue()
        self.files = []
        self.output_path = None
        self.running = False

        root = self.root = tk.Tk()
        root.title("SAP Equipment Extractor  -  Tag / Manufacturer / Model / Serial / Part No.")
        root.geometry("900x640")
        root.minsize(760, 520)
        pad = dict(padx=8, pady=4)

        # --- step 1: PDFs ------------------------------------------------------
        f1 = ttk.LabelFrame(root, text=" 1.  Upload PDF files ")
        f1.pack(fill="both", expand=True, **pad)
        bar = ttk.Frame(f1)
        bar.pack(fill="x", padx=6, pady=4)
        ttk.Button(bar, text="Add PDF file(s)...", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Add a whole folder...", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(bar, text="Remove selected", command=self.remove_selected).pack(side="left")
        ttk.Button(bar, text="Clear list", command=self.clear_files).pack(side="left", padx=6)
        self.count_lbl = ttk.Label(bar, text="No files added yet")
        self.count_lbl.pack(side="right")
        lf = ttk.Frame(f1)
        lf.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.listbox = tk.Listbox(lf, selectmode="extended", height=8)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # --- step 2: options -----------------------------------------------------
        f2 = ttk.LabelFrame(root, text=" 2.  Options ")
        f2.pack(fill="x", **pad)
        r1 = ttk.Frame(f2)
        r1.pack(fill="x", padx=6, pady=3)
        ttk.Label(r1, text="Tag list (optional, improves accuracy):", width=36).pack(side="left")
        self.tag_var = tk.StringVar(value=preset_taglist or "")
        ttk.Entry(r1, textvariable=self.tag_var, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(r1, text="Upload tag list...", command=self.pick_taglist).pack(side="left", padx=4)
        ttk.Button(r1, text="X", width=3, command=lambda: self.tag_var.set("")).pack(side="left")
        r2 = ttk.Frame(f2)
        r2.pack(fill="x", padx=6, pady=3)
        ttk.Label(r2, text="Reading mode:", width=36).pack(side="left")
        self.mode_var = tk.StringVar(value=mode)
        for val, txt in (("fast", "Fast"), ("normal", "Normal (recommended)"), ("thorough", "Thorough (bad scans)")):
            ttk.Radiobutton(r2, text=txt, value=val, variable=self.mode_var).pack(side="left", padx=6)
        r3 = ttk.Frame(f2)
        r3.pack(fill="x", padx=6, pady=3)
        ttk.Label(r3, text="Save Excel result as:", width=36).pack(side="left")
        self.out_var = tk.StringVar()
        ttk.Entry(r3, textvariable=self.out_var, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(r3, text="Change...", command=self.pick_output).pack(side="left", padx=4)

        # --- step 3: run -------------------------------------------------------------
        f3 = ttk.LabelFrame(root, text=" 3.  Extract ")
        f3.pack(fill="both", expand=True, **pad)
        r4 = ttk.Frame(f3)
        r4.pack(fill="x", padx=6, pady=4)
        self.start_btn = ttk.Button(r4, text="START EXTRACTION", command=self.start)
        self.start_btn.pack(side="left")
        self.open_btn = ttk.Button(r4, text="Open Excel result", command=lambda: open_file(self.output_path),
                                   state="disabled")
        self.open_btn.pack(side="left", padx=6)
        self.prog = ttk.Progressbar(r4, mode="determinate")
        self.prog.pack(side="left", fill="x", expand=True, padx=6)
        self.prog_lbl = ttk.Label(r4, text="")
        self.prog_lbl.pack(side="left")
        self.logbox = tk.Text(f3, height=9, state="disabled", wrap="word")
        self.logbox.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        for p in collect_pdfs(list(preset_files)):
            self._add(p)
        self.log("Add your PDF files (or a folder), then press START EXTRACTION.")
        if not tesseract_ready():
            self.log("!! Tesseract OCR was not found - scanned PDFs cannot be read until it is installed.")
            if messagebox.askyesno("Tesseract OCR not found",
                                   "Tesseract OCR is needed to read scanned PDFs but was not found.\n\n"
                                   "Do you want to locate tesseract.exe yourself?\n"
                                   "(Choose No if you have not installed it yet.)"):
                path = filedialog.askopenfilename(title="Locate tesseract.exe",
                                                  filetypes=[("tesseract", "tesseract*"), ("All files", "*.*")])
                if path and tesseract_ready(path):
                    self.tesseract = path
                    self.log("Tesseract found: " + path)
        self.tesseract = getattr(self, "tesseract", None)
        root.after(200, self.poll)

    # ---------- file list ----------
    def _add(self, path):
        path = os.path.abspath(path)
        if path not in self.files:
            self.files.append(path)
            self.listbox.insert("end", path)
        n = len(self.files)
        self.count_lbl.config(text=f"{n} PDF file(s) added")
        if not self.out_var.get() and self.files:
            self.out_var.set(os.path.join(os.path.dirname(self.files[0]),
                                          f"SAP_Equipment_Extract_{datetime.datetime.now():%Y%m%d_%H%M}.xlsx"))

    def add_files(self):
        for p in self.fd.askopenfilenames(title="Upload PDF file(s)", filetypes=[("PDF files", "*.pdf *.PDF")]):
            self._add(p)

    def add_folder(self):
        d = self.fd.askdirectory(title="Choose a folder with PDF files")
        if not d:
            return
        sub = self.mb.askyesno("Sub-folders", "Also include PDFs inside sub-folders?")
        found = collect_pdfs([d], recursive=sub)
        for p in found:
            self._add(p)
        self.log(f"Folder added: {d}  ({len(found)} PDF file(s))")

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.files.pop(i)
            self.listbox.delete(i)
        self.count_lbl.config(text=f"{len(self.files)} PDF file(s) added" if self.files else "No files added yet")

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.count_lbl.config(text="No files added yet")

    def pick_taglist(self):
        p = self.fd.askopenfilename(title="Upload tag list",
                                    filetypes=[("Excel / CSV / Text", "*.xlsx *.xlsm *.csv *.txt")])
        if p:
            self.tag_var.set(p)

    def pick_output(self):
        p = self.fd.asksaveasfilename(title="Save Excel result as", defaultextension=".xlsx",
                                      filetypes=[("Excel", "*.xlsx")],
                                      initialfile=os.path.basename(self.out_var.get() or "SAP_Equipment_Extract.xlsx"))
        if p:
            self.out_var.set(p)

    # ---------- run ----------
    def log(self, msg):
        self.logbox.configure(state="normal")
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    def start(self):
        if self.running:
            return
        if not self.files:
            self.mb.showwarning("No files", "Please add at least one PDF file first.")
            return
        out = self.out_var.get()
        if not out:
            self.pick_output()
            out = self.out_var.get()
            if not out:
                return
        self.running = True
        self.start_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.prog["value"] = 0
        self.log(f"\nStarting: {len(self.files)} file(s) ...")
        args = dict(inputs=list(self.files), output=out, taglist=self.tag_var.get() or None,
                    mode=self.mode_var.get(), tesseract=self.tesseract)
        self.threading.Thread(target=self._worker, args=(args,), daemon=True).start()

    def _worker(self, args):
        q = self.queue
        try:
            out = run(**args, log=lambda m: q.put(("log", m)),
                      progress=lambda d, t: q.put(("prog", (d, t))))
            q.put(("done", out))
        except PermissionError:
            q.put(("error", "Could not save the Excel file.\nIs it open in Excel? Close it and press START again."))
        except Exception as e:
            q.put(("error", f"{type(e).__name__}: {e}\n\n{traceback.format_exc()[-1500:]}"))

    def poll(self):
        try:
            while True:
                kind, val = self.queue.get_nowait()
                if kind == "log":
                    self.log(val)
                elif kind == "prog":
                    d, t = val
                    self.prog["maximum"] = t
                    self.prog["value"] = d
                    self.prog_lbl.config(text=f"{d}/{t} pages")
                elif kind == "done":
                    self.running = False
                    self.start_btn.config(state="normal")
                    if val:
                        self.output_path = val
                        self.open_btn.config(state="normal")
                        if self.mb.askyesno("Finished", f"Done!\n\nSaved to:\n{val}\n\nOpen the Excel file now?"):
                            open_file(val)
                    else:
                        self.mb.showwarning("Nothing to read", "No readable PDF pages were found.")
                elif kind == "error":
                    self.running = False
                    self.start_btn.config(state="normal")
                    self.log("!! " + val.splitlines()[0])
                    self.mb.showerror("Error", val)
        except Exception:
            pass
        self.root.after(200, self.poll)

    def mainloop(self):
        self.root.mainloop()


def colab_upload_flow(mode="normal"):
    """Google Colab: upload button for PDFs (and optional tag list), then download the Excel result."""
    from google.colab import files
    import zipfile
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    work = os.path.abspath(f"sap_extract_{stamp}")
    os.makedirs(work, exist_ok=True)

    print("STEP 1: click 'Choose Files' and upload your PDF files (you can select many at once,"
          " or upload one .zip containing PDFs).")
    up = files.upload()
    pdfs = []
    for name, data in up.items():
        dest = os.path.join(work, os.path.basename(name))
        with open(dest, "wb") as fh:
            fh.write(data)
        try:
            os.remove(name)          # Colab also saves a copy in the current folder
        except OSError:
            pass
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(dest) as z:
                for m in z.namelist():
                    if m.lower().endswith(".pdf") and not m.startswith("__MACOSX"):
                        out = os.path.join(work, os.path.basename(m))
                        with z.open(m) as src, open(out, "wb") as dst:
                            dst.write(src.read())
                        pdfs.append(out)
        elif name.lower().endswith(".pdf"):
            pdfs.append(dest)
    if not pdfs:
        print("No PDF files were uploaded. Run the cell again.")
        return None
    print(f"\n{len(pdfs)} PDF file(s) uploaded.")

    taglist = None
    ans = input("\nSTEP 2 (optional): upload a TAG LIST Excel/CSV to improve accuracy? (y/n): ").strip().lower()
    if ans.startswith("y"):
        tl = files.upload()
        for name, data in tl.items():
            taglist = os.path.join(work, "taglist_" + os.path.basename(name))
            with open(taglist, "wb") as fh:
                fh.write(data)
            try:
                os.remove(name)
            except OSError:
                pass
            break

    out = os.path.join(work, f"SAP_Equipment_Extract_{stamp}.xlsx")
    print("\nSTEP 3: reading the PDFs ...")
    res = run(pdfs, out, taglist, mode=mode)
    if res:
        print("\nDownloading the Excel result ...")
        files.download(res)
    return res


def ask_paths_console():
    """Fallback when no window can be opened and we are not in Colab."""
    print("Enter PDF files or folders (one per line, empty line to finish):")
    items = []
    while True:
        line = input("> ").strip().strip('"')
        if not line:
            break
        items.append(line)
    return items


def main():
    global _STANDALONE
    _STANDALONE = not IN_COLAB
    ap = argparse.ArgumentParser(description="Extract Tag / Manufacturer / Model / Serial / Part No. from PDFs "
                                             "for SAP equipment templates. Opens an upload window.")
    ap.add_argument("inputs", nargs="*", help="(optional) PDFs/folders to pre-load into the window")
    ap.add_argument("--taglist", help="(optional) tag list to pre-load")
    ap.add_argument("--no-window", action="store_true",
                    help="advanced: run without the window using the paths given on the command line")
    ap.add_argument("-o", "--output", help="output .xlsx (only with --no-window)")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--fast", action="store_true")
    ap.add_argument("--thorough", action="store_true")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--workers", type=int)
    ap.add_argument("--tesseract", help="path to tesseract executable")
    a, _unknown = ap.parse_known_args()   # ignore extra args added by Jupyter / Colab (e.g. -f)
    a.inputs = [x for x in a.inputs if not x.endswith(".json")]   # Jupyter kernel file
    mode = "fast" if a.fast else ("thorough" if a.thorough else "normal")
    if IN_COLAB:
        colab_upload_flow(mode)
        return
    if a.no_window:
        if not a.inputs:
            ap.error("--no-window needs PDF files or folders")
        run(a.inputs, a.output, a.taglist, a.recursive, mode, a.dpi, a.workers, a.tesseract)
        return
    try:
        win = UploadWindow(a.inputs, a.taglist, mode)
    except Exception as e:      # no display / tkinter missing -> console mode
        print(f"(Could not open the upload window: {e})")
        items = a.inputs or ask_paths_console()
        if items:
            run(items, a.output, a.taglist, True, mode, a.dpi, a.workers, a.tesseract)
        return
    win.mainloop()


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()   # needed on Windows (also if turned into an .exe)
    main()
