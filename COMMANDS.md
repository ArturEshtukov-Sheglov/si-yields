# Commands

All the commands needed to run si-yields, for Windows, Linux, macOS and Google Colab.
**Русская версия:** [COMMANDS_RU.md](COMMANDS_RU.md)

In every case the result is the same: `result/yields.xlsx` and `result/yields.csv`.


## 1. Install once

Python 3.9 or newer is needed (https://www.python.org; on Linux and macOS it is usually already
installed as `python3`).

| system | command (in the folder si-yields) |
|---|---|
| Windows | `python -m pip install -r requirements.txt` (or `py -m pip ...`) |
| Linux / macOS | `python3 -m pip install -r requirements.txt` |

If pip refuses to install into the system Python ("externally-managed-environment", recent
Linux and macOS), use a virtual environment:

    python3 -m venv .venv
    source .venv/bin/activate          # Windows: .venv\Scripts\activate
    python -m pip install -r requirements.txt

(activate it again in every new terminal before running the script).


## 2. Run on every PDF in data/

Put the SI PDF files into the folder `data/`, then:

| system | double-click | or in a terminal (in the folder si-yields) |
|---|---|---|
| Windows | `run.bat` | `python extract_yields.py` |
| Linux | — | `bash run.sh` (or `./run.sh`) |
| macOS | `run.command` | `bash run.sh` |

On macOS the first double-click on `run.command` may be blocked ("cannot be opened because it is
from an unidentified developer"): right-click it, choose **Open**, then **Open** again. If it does
not start at all, run once in Terminal: `chmod +x run.sh run.command`.


## 3. Other files or another result name

Not needed for files that are already in `data/`: for them section 2 is enough. File names are
given only for PDF files outside `data/`, or to process only some of them.

    python extract_yields.py SI_1.pdf SI_2.pdf             # Windows
    python3 extract_yields.py SI_1.pdf SI_2.pdf            # Linux / macOS
    python3 extract_yields.py path/to/*.pdf -o yields_2026

Without `-o`, the result of this form is written to `yields.xlsx` and `yields.csv` in the folder
where the command is run.

    python extract_yields.py --help                        # all options


## 4. Google Colab (no installation)

Open `si_yields_Colab.ipynb` in Colab (on GitHub: open the file and click "Open in Colab", or in
Colab: File → Open notebook → GitHub → paste the repository address). Run the three cells:

1. installs the packages and downloads the script;
2. asks for the PDF files (puts them into `data/`);
3. runs the script, shows the first rows and downloads `yields.xlsx`.

The PDF files are uploaded to Google's servers: for unpublished SI, run the script on your own
computer.


## 5. Check that it works

After a run, the screen shows one line per file, for example

    SI_1.pdf: 43 compounds with a yield, 37 with a name
    Written: .../result/yields.xlsx and .../result/yields.csv (43 rows)

Open `result/yields.xlsx` and compare a few rows with the `context` column and the PDF.
