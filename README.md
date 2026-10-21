# si-yields

Compound labels, names and yields from the Supporting Information (SI) of chemistry papers.

The script reads SI PDF files and writes a table with one row per compound:

| label | name | yield_percent |
|---|---|---|
| 3c | (5R,6R,9R)-4-methyl-6-(naphthalen-2-yl)-2,7,9-triphenyl-9-vinyl-2,3,7-triazaspiro[4.4]non-3-ene-1,8-dione | 66 |
| 26 | | 81 |

plus columns that help to check the result (file, note, page, context).

Author: Artur V. Eshtukov-Shcheglov. The code was written with the assistance of Claude Opus 5.5
(Anthropic) and was reviewed and tested by the author. Licence: MIT (see LICENSE).

**Русская версия:** [README_RU.md](README_RU.md)


## What it finds

A yield is read when it is written in the text of the SI in one of the usual forms, for example

    Colorless oil, (147.9 mg, 79% yield)
    White solid, m.p. 166-168 °C. 36.1 mg, 66% yield, >20:1 dr, 88% ee.
    ... to afford 3a in 72% yield (1.04 g).
    ... to afford, in order of elution, 17 (795 mg, 2.21 mmol, 45 %)
    Yield: 35% (91.8 mg)

A percentage counts as a yield only when a mass or an amount (mg, g, mmol) or the word "yield"
stands next to it. Percentages that are not yields are skipped: ee, de, conversion, purity,
eluent composition ("10-20% EtOAc in hexane"), reagents ("30% H2O2", "10% Pd/C"), table cells
and the text of reaction schemes.

The **label** is taken from the heading of the compound ("... propanoate (3c)", "3a: white
solid", "Compound 6a: ...") or from the sentence with the yield ("to afford 3a in 72% yield").
The **name** is the text before the label in the heading; it is left empty when the SI gives only
the label or when the name cannot be separated from the text around it.

If a compound has more than one yield (another set of conditions, a gram-scale run, a second
procedure), the first ordinary yield is given in `yield_percent` and the others are listed in
`note`.


## Requirements

Python 3.9 or newer and two packages:

    python -m pip install -r requirements.txt

(`pymupdf` reads the PDF files, `openpyxl` writes the Excel table; without `openpyxl` only the
CSV file is written.)


## How to run

**Simplest way:** put the SI PDF files into the folder `data/` and run

    python extract_yields.py

(Windows: double-click `run.bat`; Linux: `bash run.sh`; macOS: double-click `run.command`;
no installation: `si_yields_Colab.ipynb` in Google Colab). All commands for every system:
[COMMANDS.md](COMMANDS.md). The result is written to `result/yields.xlsx` and
`result/yields.csv`. Files in `data/` that are not PDF are listed on screen and skipped
(save SI in Word format as PDF first).

**Other files or another result name:**

    python extract_yields.py SI_1.pdf SI_2.pdf
    python extract_yields.py path/to/*.pdf -o yields_2026

Here the result is written to `yields.xlsx` and `yields.csv` (or the name given after `-o`) in the
folder where the script is run.

A table with the same name is overwritten; rename it or use `-o` to keep earlier results.

On screen, for each file: the number of compounds with a yield and how many of them have a name.
A file that cannot be read is reported and skipped; the other files are processed.


## Output columns

| column | meaning |
|---|---|
| file | SI file name |
| label | compound label as written in the SI (3a, S10, (Sa)-L15, 4z') |
| name | compound name, "" = not found |
| yield_percent | yield in % |
| note | what to check: NMR yield, gram scale / scale-up, over several steps, combined yield, brsm, "below the value given" (yield written as "<5%"), other yields of the same compound with their pages, "label taken from the sentence, no heading", "no name found" |
| page | PDF page of the yield |
| context | the text around the yield, to check the row without opening the PDF |


## IMPORTANT: check the result

The rules are simple text patterns. Check the rows against the `context` column, and look at
the PDF for rows with a `note`.

Test on about 200 SI files of organic chemistry papers (about 7,500 compounds with a yield):

- in a random check of 100 rows by hand, the label and the yield were right in 96; the wrong
  rows were a reagent percentage taken for a yield ("KH (290 mg, 30%, ...)"), a yield in the text
  of a scheme, a prime lost from a label (3c instead of 3c') and one label that could not be
  confirmed;
- a name was found for about 80% of the rows; about one name in ten starts with stray text from
  the page (a page number, a spectrum caption, a word of the previous sentence) that must be
  removed by hand;
- of the compounds that have an 1H NMR description in these SI, about 88% got a yield. Most of
  the rest have no yield in the SI at all (the yields are given only in the schemes of the paper).


## Limitations

- Yields given only in a table or a scheme (in the SI or in the paper) are not read.
- A yield is assigned to the label of the nearest heading or sentence before it; if one paragraph
  describes several compounds, check that the label is right.
- Labels written only in a structure drawing are not found.
- The text is taken from the PDF as it is: scanned SI (images without text) give nothing, and
  SI in Word format must first be saved as PDF.
- Only percentages are read; a yield written as a fraction or only as a mass is not.
