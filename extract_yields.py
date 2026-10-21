"""Compound labels, names and yields from Supporting Information (SI) PDF files.

For every yield written in the experimental part of an SI ("(147.9 mg, 79% yield)",
"36.1 mg, 66% yield", "afforded 3a in 72% yield", "17 (795 mg, 2.21 mmol, 45 %)") the script
finds the compound label it belongs to (3a, S10, (Sa)-L15 ...), the compound name where the SI
gives one, and writes one row per compound to an Excel table and a CSV file.

    python extract_yields.py                      (every PDF in data/ -> result/yields.xlsx, .csv)
    python extract_yields.py SI_1.pdf [SI_2.pdf ...] [-o yields]

The result must be checked: see README.md (columns "note" and "context").

Author: Artur V. Eshtukov-Shcheglov. Code written with the assistance of Claude Opus 5.5
(Anthropic); reviewed and tested by the author. Licence: MIT.
"""

import argparse
import csv
import os
import re
import sys

try:
    import pymupdf
except ImportError:
    sys.exit("The package pymupdf is missing. Install it:  python -m pip install pymupdf openpyxl")


# ===== Reading the PDF =====

# lines that the PDF puts into the running text: download stamps, page numbers ("S12", "25")
NOISE_LINE = re.compile(r"^\s*(?:Downloaded from .*|S?\s?\d{1,3}|Page \d+.*)\s*$")


def pdf_text(path):
    """Text of the PDF with noise lines removed; also the start offset of every page."""
    parts, starts, pos = [], [], 0
    with pymupdf.open(path) as doc:
        for page in doc:
            lines = [ln for ln in page.get_text().splitlines() if not NOISE_LINE.match(ln)]
            text = "\n".join(lines) + "\n"
            starts.append(pos)
            parts.append(text)
            pos += len(text)
    text = "".join(parts)
    text = text.replace("​", "").replace(" ", " ")
    return text, starts


def page_of(offset, starts):
    n = 0
    for i, s in enumerate(starts):
        if s <= offset:
            n = i
    return n + 1


# ===== Compound labels =====

# 3a, 3ag, 1t, 22, S10, 4b', SI-7, (Sa)-L15, [D1]-1e
LABEL = (r"(?:[\(\[][A-Za-z0-9,+\-]{1,8}[\)\]]-)?"   # optional stereo/isotope prefix: (Sa)-, [D1]-
         r"(?:[A-Z]{1,3}-?)?"                      # optional letters: S, L, SI-
         r"\d{1,3}[a-z]{0,3}"                      # number + letters: 3, 3a, 3ag
         r"['′″’]*")                               # primes
# the whole parenthesis is a label: "... propanoate (3c)"
LABEL_PAREN = re.compile(r"\((" + LABEL + r")\)")
# label at the start of a line, followed by ":" or ".": "3a: white solid", "3a. Methyl ..."
LABEL_LINE = re.compile(r"(?m)^[ \t]*(" + LABEL + r")[ \t]*[:.][ \t]")
# label right before the yield: "afford 3a in 72% yield", "17 (795 mg", "title compound 29 (450 mg"
VERB = r"(?:afford(?:ed|s)?|gave|give[sn]?|provid(?:ed|es|e)|furnish(?:ed|es)?|yield(?:ed|s)?|obtain(?:ed)?|deliver(?:ed|s)?|isolat(?:ed|e))"
LABEL_BEFORE_YIELD = re.compile(
    r"(?<![\w.,\-:=/])(" + LABEL + r")\s*(?:\(\s*)?"
    r"(?:(?:was|is|were)\s+(?:obtained|isolated|prepared|synthesi[sz]ed|afforded|formed)\s*)?"
    r"(?:as\s+(?:an?\s+)?[^()]{0,40}?)?\s*(?:in\s*)?$")
# "2-(1-Phenylvinyl)aniline S2b:", "Alkoxyamine 1a: ..." - a label with letters, then ":"
LABEL_COLON = re.compile(r"(?<=\s)((?:[A-Z]{1,3}-?)?\d{1,3}[a-z]{1,3}['′]*|[A-Z]{1,3}-?\d{1,3}['′]*)\s?:(?!\d)")

SECTION_WORDS = re.compile(r"\s*(?:Synthesis|Preparation|General|Experimental|Characteri[sz]ation|Procedure|"
                           r"Materials|Optimi[sz]ation|Mechanistic|Control|Scale|Gram|Large|Results|"
                           r"Reaction|Typical|NMR|X-ray|Crystal|Computational|DFT|References|Supplementary)\b")
NOT_LABEL = re.compile(r"^(?:\d{4,}|[A-Z]{1,3}|(?:GP|GM|GC|TS|SM|ESI|APCI|EI|MS|CCDC|Ref|Eq|Fig|Table|Entry|Method|Step)-?\d.*|(?:CO|SO|NO|CH|CF|CD|H|O|N|C)\d)$", re.I)


def clean_label(lab):
    return lab.replace("′", "'").replace("’", "'").replace("″", "''").strip()


# ===== Yields =====

PERCENT = re.compile(r"(?<![\d.\-–])(\d{1,3}(?:[.,]\d)?)\s?%")
AMOUNT = re.compile(r"\d+(?:[.,]\d+)?\s?(?:mg|g|mmol|μmol|µmol|umol|mol|kg)\b")
YIELD_WORD = re.compile(r"yield", re.I)
# the percentage is not a yield when followed or preceded by these words
# a reagent: "30% H2O2", "10% Pd/C", "37% HCl" (case-sensitive)
REAGENT_AFTER = re.compile(r"^\s*(?:[A-Z][a-z]?(?:\d|[A-Z])|Pd|Pt|Rh|Ru|Ni)")
NOT_YIELD_AFTER = re.compile(r"^\s*(?:aqueous|ee|e\.e\.|de|d\.e\.|conv|purity|dispersion|solution|w/w|wt|aq|in\s+(?:mineral|oil|water|hexane|pentane|heptane|petroleum|EtOAc|DCM|CH2Cl2|MeOH)|[A-Z][a-z]?OAc|EtOAc|ethyl|hexane|petroleum|MeOH|methanol|DCM|CH2Cl2|acetone|Et2O|ether|v/v|of\s+the)", re.I)
NOT_YIELD_BEFORE = re.compile(r"(?:conversion|purity|ee|>|≥|ca\.|HPLC|purity of|dispersion)\s*(?:of\s*)?$", re.I)


def is_yield(text, m):
    """Is the percentage at match m a yield? Returns (yes, note)."""
    after = text[m.end():m.end() + 25]
    before = text[max(0, m.start() - 60):m.start()]
    if NOT_YIELD_AFTER.match(after) or REAGENT_AFTER.match(after) or NOT_YIELD_BEFORE.search(before[-20:]):
        return False, ""
    # a range "10-20%" (eluent) is excluded by the look-behind of PERCENT; a ratio "1:1" is not a %
    clause = re.split(r"[;\n]\s*\n|\.\s+[A-Z]", before)[-1]          # same sentence
    line = text[text.rfind("\n", 0, m.start()) + 1:text.find("\n", m.end())]
    if re.fullmatch(r"\s*\d{1,3}(?:[.,]\d)?\s?%\s*(?:\[\w\])?\s*", line):      # a table cell "91%"
        return False, ""
    prev_lines = [x for x in text[max(0, m.start() - 300):m.start()].split("\n")[:-1] if x.strip()][-4:]
    if len(line.strip()) < 15 and len(prev_lines) == 4 and all(len(x.strip()) < 15 for x in prev_lines):
        return False, ""                                                     # text of a scheme or table
    near_amount = AMOUNT.search(clause[-45:]) is not None or AMOUNT.match(after.lstrip(" (")) is not None
    near_word = YIELD_WORD.search(before[-25:] + after[:22]) is not None
    if not (near_amount or near_word):
        return False, ""
    note = []
    if re.search(r"NMR\s+yield|yield\s*\(?\s*NMR|by\s+(?:1H\s+)?NMR", before[-40:] + after[:25], re.I):
        note.append("NMR yield")
    if re.search(r"brsm|based on recovered", before[-40:] + after[:40], re.I):
        note.append("brsm")
    if re.search(r"<\s*$", before[-3:]):
        note.append("below the value given")
    if re.search(r"combined", before[-30:] + after[:15], re.I):
        note.append("combined yield")
    if re.search(r"(?:over|for|in)\s+(?:two|three|four|five|six|\d+)\s+steps", after[:40], re.I):
        note.append("over several steps")
    return True, ", ".join(note)


# ===== Compound names =====

STOP_WORDS = set("""the a an and or was were is are be been with without to in into for by as from at on
then after before under over yield yields afford afforded gave give given provided obtained product
compound general procedure synthesis preparation prepared of following according using solid oil
white yellow colorless colourless brown red orange pale light dark found calcd calculated hrms nmr
eluent m.p mp rf step reaction mixture residue purified column chromatography silica min minor major
mhz spectrum spectra products data analytical""".split())
DRAWING_TOKEN = re.compile(r"^(?:[A-Z][a-z]?|Me|Et|Pr|Bu|Ph|Bn|Bz|Ts|Ms|Tf|Ac|Boc|Cbz|TBS|TIPS|HO|OH|NH|NH2|OMe|CO2Me|CO2Et|R\d?|Ar|X)$")
CHEM_PART = re.compile(r"(?:yl|one|ol|ate|ine|ide|ane|ene|yne|ile|oic|ium|ose|ole|azole|amine|amide|acid|ester|ether|oxy|thio|carb|benz|phen|pyr|ind|quin|furan|cyclo|methyl|ethyl|propyl|butyl)", re.I)


def name_before(text, end):
    """Compound name ending at offset `end` (right before "(3a)"), or ""."""
    chunk = text[max(0, end - 400):end]
    tokens = re.split(r"(\s+)", chunk)
    words = []
    for tok in reversed(tokens):
        if not tok or tok.isspace():
            if words:
                words.append(" ")
            continue
        low = tok.lower().strip(",;()")
        if (tok.endswith((".", ":", ";")) and not re.search(r"[A-Za-z]\.$", tok[-3:]) or low in STOP_WORDS
                or "%" in tok or re.fullmatch(r"[\d.,]+", tok) or tok.startswith(("δ", "(ESI", "Hz"))):
            break
        if (tok.endswith((".", ":")) or re.search(r"[.\]]\)$|\.\[\d+\]|\.\d+$|MHz|CDCl3\)|DMSO", tok)
                or re.fullmatch(r"(?:SI|S)-?\d+|\(\d+|\(\w{0,3}\)|N-NH\w*", tok)):
            break
        if re.fullmatch(LABEL, tok) or re.fullmatch(r"\(\W*\)", tok):      # another label, "()"
            break
        if DRAWING_TOKEN.match(tok):          # atom labels of a structure drawing: "N O Me HO"
            break
        words.append(tok)
    name = "".join(reversed(words)).strip()
    # a line-break hyphen inside a name: "triazaspiro[4.4]non- 3-ene" -> "triazaspiro[4.4]non-3-ene"
    name = re.sub(r"-\s+(?=[\d\[(A-Za-z])", "-", name)
    name = re.sub(r"\s+", " ", name)
    name = re.sub(r"^(?:Supporting Information\s+)?[•·▪\-–]?\s*", "", name)
    name = re.sub(r"^(?:Compound|Product|Title compound)\s+", "", name, flags=re.I)
    name = re.sub(r"\s*\(" + LABEL + r"\)\s*$", "", name)          # "... methanone (3g)" before "3g:"
    if len(name) < 6 or not CHEM_PART.search(name) or name.count("(") != name.count(")"):
        return ""
    return name


# ===== Main search =====

def find_label(text, start, stop):
    """Label for a yield at `stop`, searching back to `start` (end of the previous yield).
    Returns (label, how, label offset, end of the label match)."""
    window = text[start:stop]
    # text before the yield phrase: cut at the "(" of "(75 mg, 86%" or at the amount of "36.1 mg, 66%"
    tail = window[-90:]
    op = tail.rfind("(")
    if op != -1 and ")" not in tail[op:]:
        pre = tail[:op]
    else:
        am = list(AMOUNT.finditer(tail))
        pre = tail[:am[0].start()] if am else tail
    pre = re.sub(r"-\s*\n\s*", "", pre)          # "af-\nforded"
    m = LABEL_BEFORE_YIELD.search(pre.rstrip(" ,\n"))
    cand = []
    for mm in LABEL_PAREN.finditer(window):
        if not re.search(r"[A-Za-z]{3}", window[max(0, mm.start() - 15):mm.start()]):
            continue                     # "0.2 (100h)" in a table: no name before the parenthesis
        cand.append((mm.start(), clean_label(mm.group(1)), "heading", start + mm.start(), start + mm.start()))
    for mm in LABEL_LINE.finditer(window):
        if SECTION_WORDS.match(window[mm.end():mm.end() + 40]):
            continue                     # "2a. Synthesis of starting materials": a section number
        cand.append((mm.start(), clean_label(mm.group(1)), "heading", start + mm.start(), None))
    for mm in LABEL_COLON.finditer(window):
        cand.append((mm.start(), clean_label(mm.group(1)), "heading", start + mm.start(), start + mm.start()))
    cand = [c for c in cand if not NOT_LABEL.match(c[1])]
    # a label written right before the yield wins ("afford 3a in 72% yield")
    if m and not NOT_LABEL.match(m.group(1)):
        lab = clean_label(m.group(1))
        heading = [c for c in cand if c[1] == lab]
        if heading:
            c = heading[-1]
            return lab, "heading", c[3], c[4]
        return lab, "in text", None, None
    if cand:
        c = max(cand, key=lambda c: c[0])
        return c[1], c[2], c[3], c[4]
    return "", "", None, None


def extract(path):
    text, starts = pdf_text(path)
    rows, order = {}, []
    prev_end = 0
    last_label = ""
    for m in PERCENT.finditer(text):
        ok, note = is_yield(text, m)
        if not ok:
            continue
        value = m.group(1).replace(",", ".")
        lab, how, lab_pos, name_end = find_label(text, max(prev_end, m.start() - 1500), m.start())
        if not lab and last_label and m.start() - prev_end < 200:
            lab, how = last_label, "same paragraph"          # second yield of the same compound
        prev_end = m.end()
        if not lab:
            continue
        last_label = lab
        name = name_before(text, name_end) if name_end is not None else ""
        before = text[max(0, m.start() - 700):m.start()]
        if re.search(r"gram[\s-]?scale|scale[\s-]?up|large[\s-]?scale|\d+(?:\.\d+)?\s?mmol\s+scale", before, re.I):
            note = ", ".join(x for x in (note, "gram scale / scale-up") if x)
        context = re.sub(r"\s+", " ", text[max(0, m.start() - 110):m.end() + 30]).strip()
        context = "…" + context.split(" ", 1)[-1].rsplit(" ", 1)[0] + "…"   # whole words only
        rec = dict(value=value, note=note, page=page_of(m.start(), starts), context=context, how=how, pos=m.start())
        if lab not in rows:
            rows[lab] = dict(file=os.path.basename(path), label=lab, name=name, yields=[rec])
            order.append(lab)
        else:
            if name and not rows[lab]["name"]:
                rows[lab]["name"] = name
            rows[lab]["yields"].append(rec)
    out = []
    for lab in order:
        r = rows[lab]
        main = next((y for y in r["yields"] if not y["note"]), r["yields"][0])
        others = [y for y in r["yields"] if y is not main]
        notes = [main["note"]] if main["note"] else []
        if others:
            notes.append("other yields: " + "; ".join(
                f'{y["value"]}%' + (f' ({y["note"]})' if y["note"] else "") + f' p.{y["page"]}' for y in others))
        if main["how"] == "in text":
            notes.append("label taken from the sentence, no heading")
        if not r["name"]:
            notes.append("no name found")
        out.append(dict(file=r["file"], label=lab, name=r["name"], yield_percent=float(main["value"]),
                        note="; ".join(notes), page=main["page"], context=main["context"], pos=main["pos"]))
    return out


# ===== Output =====

COLUMNS = ["file", "label", "name", "yield_percent", "note", "page", "context"]


def write(rows, stem):
    with open(stem + ".csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    try:
        import openpyxl
        from openpyxl.styles import Font
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    except ImportError:
        print("openpyxl is not installed: only the CSV file was written.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "yields"
    ws.append(COLUMNS)
    for r in rows:
        ws.append([ILLEGAL_CHARACTERS_RE.sub("", str(r[c])) if isinstance(r[c], str) else r[c] for c in COLUMNS])
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, width in zip("ABCDEFG", (22, 10, 60, 13, 40, 6, 90)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(stem + ".xlsx")


HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
RESULT = os.path.join(HERE, "result")


def main():
    ap = argparse.ArgumentParser(description="Compound labels, names and yields from SI PDF files. "
                                 "Without file names: every PDF in the folder data/, result in result/.")
    ap.add_argument("pdf", nargs="*", help="SI PDF file(s); default: every PDF in data/")
    ap.add_argument("-o", default=None, help="name of the result files without extension "
                    "(default: result/yields when data/ is used, otherwise yields)")
    args, _ = ap.parse_known_args()      # ignore extra arguments (e.g. from a notebook kernel)
    pdfs = args.pdf
    out = args.o
    if not pdfs:
        os.makedirs(DATA, exist_ok=True)
        pdfs = sorted(os.path.join(DATA, f) for f in os.listdir(DATA) if f.lower().endswith(".pdf"))
        if not pdfs:
            sys.exit(f"No PDF files in {DATA}\nPut the SI PDF files there and run the script again.")
        others = [f for f in os.listdir(DATA) if not f.lower().endswith(".pdf") and not f.startswith(".")]
        if others:
            print("Skipped (not PDF; save as PDF first): " + ", ".join(sorted(others)))
        if out is None:
            os.makedirs(RESULT, exist_ok=True)
            out = os.path.join(RESULT, "yields")
    out = out or "yields"
    rows = []
    for p in pdfs:
        try:
            r = extract(p)
        except Exception as e:                # one bad file does not stop the others
            print(f"{os.path.basename(p)}: could not be read ({e})")
            continue
        named = sum(1 for x in r if x["name"])
        print(f"{os.path.basename(p)}: {len(r)} compounds with a yield, {named} with a name")
        rows += r
    stem = os.path.splitext(out)[0]
    write(rows, stem)
    print(f"Written: {stem}.xlsx and {stem}.csv ({len(rows)} rows)")


if __name__ == "__main__":
    main()
