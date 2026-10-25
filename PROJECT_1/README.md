# PROJECT_1 — Audio testimony analysis and merging

[Back to the repository](../README.md) · [Go to PROJECT_2](../PROJECT_2/README.md)

The first stage of the Smart Factory project: a Python and Jupyter pipeline that analyzes audio testimonies about an industrial incident, extracts information using a shared schema, and generates Excel reports.

The code is in the [script_finale 3.3 .ipynb](<script_finale 3.3 .ipynb>) notebook.

## Features

| Mode | Processing | Output |
|---|---|---|
| **A — Individual analysis** | Analyzes each audio file separately with Gemini and extracts a JSON report. | One `report_singolo_<audio_name>.xlsx` file per testimony. |
| **B — Merging** | Sends the reports from mode A to Gemini to produce a synthesis of the sources. | One `report_multiaudio.xlsx` file. |

The notebook includes functions for displaying previews and comparing individual reports with the merged report.

Mode B makes a new API call to the model using the extracted JSON reports, without uploading the audio again. Some internal comments describe a local merge; the implemented `fondi_risultati_singoli` function uses Gemini instead.

## Requirements

- Python 3 and a Jupyter environment.
- An Internet connection and a Gemini API key.
- The `google-generativeai`, `pandas`, and `openpyxl` libraries, plus `IPython` for previews.
- Locally supplied audio files to analyze.

The model specified in the code is `gemini-2.5-flash`. The audio formats listed in the configuration are `.mp3`, `.wav`, `.m4a`, `.ogg`, and `.flac`.

## Installation and launch

Run the following commands in **PowerShell**, starting from the repository root.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install jupyterlab google-generativeai pandas openpyxl
$env:GEMINI_API_KEY = "your_api_key"
cd PROJECT_1
..\.venv\Scripts\python.exe -m jupyterlab "script_finale 3.3 .ipynb"
```

If a `.venv` environment already exists, you can reuse it. In Jupyter, select the kernel belonging to the environment where you installed the libraries.

The notebook also contains a `%pip install` cell for the pipeline dependencies. If you use it, restart the kernel after installation and resume from the imports.

## Configuration

In the configuration cell, labeled **CONFIGURAZIONE**, update `FILE_AUDIO` with your files, for example:

```python
FILE_AUDIO = [
    "Elena RSPP.mp3",
    "Marco Capolinea.mp3",
    "Martina Operatrice.mp3",
]

OUTPUT_DIR = "output_report"
```

These three names are the defaults in the notebook; the audio files are not included in the repository. You can replace them with absolute paths or paths relative to the kernel's working directory. With the launch commands above, relative paths start from `PROJECT_1/`.

The API key is read from `os.environ` through `GEMINI_API_KEY`. Set it before starting Jupyter; if Jupyter is already running, restart it from the configured terminal. The notebook does not automatically load a `.env` file.

## Execution

1. Run the imports, configuration, prompts, and function definitions in notebook order.
2. Check that the configuration finds every file in `FILE_AUDIO`.
3. Run the **ESEGUI — MODALITÀ A** cell (run mode A) to generate previews and separate Excel files.
4. Run the **ESEGUI — MODALITÀ B** cell (run mode B) in the same session: it uses `risultati_singoli`, created by mode A.

Follow the cell titles: some numbered references in the comments do not match the actual order. After restarting the kernel, you need to recreate the definitions and results in memory.

To display the side-by-side comparison, run a new cell after merging:

```python
mostra_confronto(risultati_singoli, dati_multi)
```

## Generated reports

With the default configuration and when launched from `PROJECT_1/`, results are written to:

```text
PROJECT_1/output_report/
├── report_singolo_Elena RSPP.xlsx
├── report_singolo_Marco Capolinea.xlsx
├── report_singolo_Martina Operatrice.xlsx
└── report_multiaudio.xlsx
```

Each file contains the worksheets `EVENTO` (incident), `LUOGHI & MEZZI` (locations and equipment), `ANALISI CAUSE` (cause analysis), `CONSEGUENZE` (consequences), `AZIONI CORRETTIVE` (corrective actions), and `PERSONE COINVOLTE` (people involved), formatted for readability.

The output directory is created automatically and ignored by Git. Running the pipeline again with the same names may overwrite previous files.

## Implementation notes

- Audio files are uploaded through the Gemini APIs; the code includes remote file deletion after processing.
- If JSON parsing fails during merging, the function returns the first individual report as a fallback: check the notebook messages before treating the result as a completed merge.
- The prompt allows inferences and proposed corrective actions. Check the information against the original testimonies.

The extension with repeated analysis, metrics, and a web interface is documented in [PROJECT_2](../PROJECT_2/README.md).
