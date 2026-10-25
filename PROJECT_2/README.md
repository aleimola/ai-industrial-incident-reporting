# PROJECT_2 — Quality analysis and the SicuraFlow web app

[Back to the repository](../README.md) · [See PROJECT_1](../PROJECT_1/README.md)

The second stage of the Smart Factory project. It extends audio testimony analysis with repeated extractions and quality metrics and introduces **SicuraFlow**, a web app for submitting incident reports, grouping reports about the same event, and exploring them through dashboards and a chatbot.

The project contains two components that can be used independently:

- [Notebook script_finale_4.4.ipynb](script_finale_4.4.ipynb): pipeline experimentation and quality analysis with Excel export.
- [Web app](Webapp/): a browser interface and FastAPI backend with its own pipeline implementation in `backend/pipeline_adapter.py`. Running the notebook is not required.

## Structure

```text
PROJECT_2/
├── README.md
├── script_finale_4.4.ipynb
└── Webapp/
    ├── server.py                  # FastAPI endpoints and server startup
    ├── index.html                 # HTML, CSS, and JavaScript interface
    ├── requirements.txt           # Web app dependencies
    ├── avvia_app.bat              # Quick launch on Windows
    ├── GUIDA_AVVIO_E_USO.md        # Guide to screens and operations
    └── backend/
        ├── __init__.py
        └── pipeline_adapter.py    # Audio analysis, reports, merging, and chatbot
```

## Notebook: available modes

| Mode | Function | Output |
|---|---|---|
| **A — Individual analysis** | Extracts a report from each audio file. | `report_singolo_<audio_name>.xlsx` |
| **B — Merging** | Merges mode A reports using Gemini. | `report_multiaudio.xlsx` |
| **C — Quality analysis** | Repeats extraction N times on the same audio file and compares the results. | `report_qualita_<audio_name>.xlsx` |

Mode B merging makes a new Gemini call using the JSON reports, even though some notebook comments state that it does not make additional API calls.

### Notebook installation and configuration

From **PowerShell**, in the repository root:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install jupyterlab google-generativeai pandas openpyxl scikit-learn scipy
$env:GEMINI_API_KEY = "your_api_key"
cd PROJECT_2
..\.venv\Scripts\python.exe -m jupyterlab script_finale_4.4.ipynb
```

You can reuse an existing `.venv` environment. The notebook also contains a dependency installation cell: restart the kernel after running it.

In the configuration cell, labeled **CONFIGURAZIONE**, update `FILE_AUDIO` with the paths to your audio files. The default files, `Elena RSPP.mp3`, `Marco Capolinea.mp3`, and `Martina Operatrice.mp3`, are not included in the repository. Relative paths are resolved from the kernel's working directory, which is `PROJECT_2/` with these launch commands.

The notebook reads `GEMINI_API_KEY` from the process environment and does not automatically load `.env`. Set the variable before starting Jupyter. The model configured in the code is `gemini-2.5-flash`.

### Running the notebook

1. Run the initial cells containing imports, configuration, prompts, and function definitions.
2. For individual and merged reports, run **ESEGUI — MODALITÀ A** (run mode A), then **ESEGUI — MODALITÀ B** (run mode B).
3. For quality analysis, run **CONFIGURAZIONE MODALITÀ C** (mode C configuration) and set `N_ESTRAZIONI` (default: `3`), HSE terms, and any reference transcripts.
4. Run **TRASCRIZIONE AUDIO** (audio transcription) if you want transcripts for each repetition.
5. Run the metric definitions, Excel export functions, and the cell defining `_runs_a_dataframes`.
6. Run the final **ESEGUI — MODALITÀ C** cell (run mode C).

Mode C uses the initial definitions and configured audio files; it does not require mode A results. Follow the cell titles and order, as some numbered references in the comments are inconsistent.

With `OUTPUT_DIR = "output_report"`, files are saved in the `output_report/` directory relative to the kernel's working directory. Quality reports include worksheets for the six report sections, `TRASCRIZIONI` (transcripts), and `ANALISI QUALITA` (quality analysis), with results for each extraction and metric summaries.

### Notebook metrics

| Metric | How it is used |
|---|---|
| Cohen's Kappa | An indicator of agreement between extractions on categorical fields, according to the implemented function. |
| Precision, Recall, and F1 | Comparison of text fields against a reference built from the extractions, using a Jaccard similarity threshold. |
| Lexical consistency | Average Jaccard similarity between pairs of extractions on selected text fields. |
| CER and WER | Character-level and word-level Levenshtein distance against a manual reference transcript. |
| DTA | Proportion of configured HSE terms present in the transcript. |
| EPR | A heuristic based on lexical overlap between the transcript and report fields. |

The notebook's **Ground Truth Proxy** uses the most frequent value for categorical fields and the longest text for text fields. This reference is derived from the model's own results: the metrics measure agreement with the proxy without certifying factual correctness.

For CER and WER, fill in `GROUND_TRUTH_TRASCRIZIONI` with manual transcripts and generate transcripts for the individual runs. Without these inputs, the values remain unavailable. DTA and EPR can also be calculated without a manual reference if generated transcripts are available. Increasing `N_ESTRAZIONI` increases the number of API calls.

## SicuraFlow web app

### Features

| Area | Main operations |
|---|---|
| Operator | Audio upload or recording, report creation and review, confirmation, and access to the operator's own submissions. |
| Incident folders | Grouping submissions about the same event, browsing testimonies, and generating a merged report. |
| Manager | Dashboard, folder filters, quality analysis, and notifications. |
| Chatbot | Questions about an incident folder and a response archive exportable as CSV. |
| Quality analysis | Comparison of repeated extractions and merges, plus indicators for the chatbot archive. |

In the backend, analyzing an audio file involves one Gemini call for transcription and analysis of the account, followed by another call to extract the structured JSON. Merging, incident matching, and the chatbot use dedicated functions in the same adapter.

### Web app installation and startup

You need Python **3.10 or later** for the syntax used by the backend, a browser, and an Internet connection for Gemini.

From **PowerShell**, in the repository root:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r PROJECT_2/Webapp/requirements.txt
cd PROJECT_2/Webapp
```

If the environment already exists, skip the first command. Create a `.env` file in `PROJECT_2/Webapp/` containing:

```dotenv
GEMINI_API_KEY=your_api_key
```

Start the server from the same directory:

```powershell
..\..\.venv\Scripts\python.exe server.py
```

Open **<http://localhost:8000>**. The health check is available at <http://localhost:8000/api/health>, and interactive API documentation is available at <http://localhost:8000/docs>.

Start the server from `Webapp/`: the code uses relative paths for `index.html`, static files, and temporary audio. To stop it, press `Ctrl+C` in the terminal.

Alternatively, on Windows, run `avvia_app.bat` from the `Webapp/` directory with your chosen Python environment active. The script runs `pip install -r requirements.txt` and `python server.py` using the commands available in the current environment.

### Demo profiles and first use

| Profile | Username | Password |
|---|---|---|
| Operator | `martina` | `password` |
| Line supervisor | `marco` | `password` |
| Safety officer (RSPP) | `elena` | `password` |
| Manager | `manager1` | `password` |

The interface lets you select demo profiles; accounts are defined in the frontend. You can also configure a personal key through **SETUP** on the opening screen: this setting is stored in the browser and sent to the server.

To try the complete workflow:

1. Select an operator and create a new submission by uploading an audio file or recording through the microphone.
2. Wait for the analysis, review the report, make any corrections, and confirm it.
3. Repeat with other testimonies about the same event to try incident folder matching.
4. Sign in as a manager, open the folder, generate a merged report, and query the chatbot.
5. Open **Analisi Qualità** (Quality Analysis) and upload an audio file to compare multiple extractions. This operation requires a separate upload.

The [startup and usage guide](Webapp/GUIDA_AVVIO_E_USO.md), written in Italian, describes the screens in more detail. In this repository checkout, the application directory is `PROJECT_2/Webapp/`.

### Data and output

- **Browser storage:** submissions, the chatbot archive, and notifications are stored in `localStorage`, within the same browser and for the same access address. There is no shared database; clearing browser data removes the local archive.
- **Temporary audio:** uploads are written to `Webapp/temp_audio/` and removed when analysis requests finish; the backend also includes deletion of remote Gemini uploads after processing.
- **Excel reports:** individual audio analysis generates files in `Webapp/output_report/`. Repeated runs using the same name may overwrite the previous file.
- **Version control:** `.env`, `temp_audio/`, `output_report/`, virtual environments, and caches are excluded from Git.

### Differences between web app and notebook metrics

The web app has its own metric implementation. The backend proxy uses the mode for categorical fields and, for text fields, the text with the highest average Jaccard similarity to the other results, known as the medoid.

In the backend's audio analysis, CER and WER compare transcripts against the first run's transcript; DTA is derived from field comparisons with the proxy, and EPR is calculated as `1 - precision`. The indicator named Kappa also uses a simplified formula. These values therefore have a different meaning from those in the notebook and are not directly comparable. Some other interface views include demo indicators.

Chatbot archive indicators describe aspects such as completeness, response length, folder coverage, and repeated questions. Evaluating correctness requires reviewing the content.

## Troubleshooting

| Problem | What to check |
|---|---|
| The notebook cannot find audio files | Check `FILE_AUDIO` and the kernel's working directory. |
| Variables or functions are undefined | Rerun the preparation cells after restarting the kernel. |
| The page does not open | Keep the server running and open `http://localhost:8000`. |
| The server cannot find `index.html` | Start it from `PROJECT_2/Webapp/`. |
| AI analysis fails | Check the key, connection, and backend response for API errors. |
| Submissions are missing | Use the same browser, address, and profile used to create them. |

SicuraFlow is a prototype with demo access and local browser storage. Generated reports require review, particularly inferred fields and corrective actions proposed by the model.
