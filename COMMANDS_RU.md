# Команды

Все команды для запуска si-yields в Windows, Linux, macOS и Google Colab.
Основной является английская версия (COMMANDS.md); это перевод для удобства.

Результат во всех случаях одинаковый: `result/yields.xlsx` и `result/yields.csv`.


## 1. Установить один раз

Нужен Python 3.9 или новее (https://www.python.org; в Linux и macOS он обычно уже установлен как
`python3`).

| система | команда (в папке si-yields) |
|---|---|
| Windows | `python -m pip install -r requirements.txt` (или `py -m pip ...`) |
| Linux / macOS | `python3 -m pip install -r requirements.txt` |

Если pip отказывается ставить пакеты в системный Python («externally-managed-environment», новые
версии Linux и macOS), используйте виртуальное окружение:

    python3 -m venv .venv
    source .venv/bin/activate          # Windows: .venv\Scripts\activate
    python -m pip install -r requirements.txt

(в каждом новом терминале его нужно снова активировать перед запуском).


## 2. Запуск на всех PDF в data/

Положите PDF-файлы SI в папку `data/`, затем:

| система | двойной щелчок | или в терминале (в папке si-yields) |
|---|---|---|
| Windows | `run.bat` | `python extract_yields.py` |
| Linux | — | `bash run.sh` (или `./run.sh`) |
| macOS | `run.command` | `bash run.sh` |

В macOS первый двойной щелчок по `run.command` может быть заблокирован («не удаётся открыть,
так как разработчик не определён»): щёлкните правой кнопкой, выберите **Открыть**, затем ещё раз
**Открыть**. Если файл не запускается совсем, один раз выполните в Терминале:
`chmod +x run.sh run.command`.


## 3. Другие файлы или другое имя результата

Для файлов, которые уже лежат в `data/`, это не нужно: для них достаточно раздела 2. Имена
файлов указываются только для PDF вне `data/` или чтобы обработать лишь некоторые из них.

    python extract_yields.py SI_1.pdf SI_2.pdf             # Windows
    python3 extract_yields.py SI_1.pdf SI_2.pdf            # Linux / macOS
    python3 extract_yields.py путь/к/*.pdf -o yields_2026

Без `-o` в этом режиме результат записывается в `yields.xlsx` и `yields.csv` в папке, из которой
запущена команда.

    python extract_yields.py --help                        # все параметры


## 4. Google Colab (без установки)

Откройте `si_yields_Colab.ipynb` в Colab (на GitHub: откройте файл и нажмите «Open in Colab»,
или в Colab: Файл → Открыть блокнот → GitHub → вставьте адрес репозитория). Выполните три ячейки:

1. ставит пакеты и скачивает скрипт;
2. просит выбрать PDF-файлы (кладёт их в `data/`);
3. запускает скрипт, показывает первые строки и скачивает `yields.xlsx`.

PDF-файлы загружаются на серверы Google: для неопубликованных SI запускайте скрипт на своём
компьютере.


## 5. Проверить, что всё работает

После запуска на экране по строке на каждый файл, например

    SI_1.pdf: 43 compounds with a yield, 37 with a name
    Written: .../result/yields.xlsx and .../result/yields.csv (43 rows)

Откройте `result/yields.xlsx` и сверьте несколько строк с колонкой `context` и с PDF.
