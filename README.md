# MMU Dataset Builder — Gemini Multimodal MCQ Creator

A high-performance Python desktop application designed to build, clean, format, and bundle multimodal Bengali Multiple-Choice Question (MCQ) datasets effortlessly using Google's Gemini Vision & Language API.

---

## 🌟 Highlights & Capabilities

- **Multimodal Question Extraction (Vision OCR)**: Paste screenshots (`Ctrl+V`) of questions from books, exam papers, or PDFs. Gemini extracts the question and options into standardized Unicode Bengali.
- **Legacy & Garbled Text Furnisher**: Have legacy Bijoy, ANSI, or corrupted OCR text (e.g. `GKwU †evgviæ wegvb f‚wg...`)? Paste it into the raw text box—Gemini automatically decodes and structures it with options.
- **Visual Diagram Manager**: Paste or browse figures, charts, tables, or geometry diagrams. Images are stored in `images/` with clean sequential IDs (`<source>_001.png`) and automatically embedded into Excel.
- **Quick-Answer Toolbar**: 1-click option selection (`[ ক ] [ খ ] [ গ ] [ ঘ ]`) with full-width responsive scaling, plus support for free-form custom answers.
- **Complete Export Pipeline**: 1-click packaging (`Ctrl+E`) generates synchronized `.xlsx`, `.csv` (UTF-8 BOM), `.jsonl` (training/eval ready), standalone responsive `.html` preview, and bundles everything into a timestamped `.zip`.
- **DPI-Aware Responsive UI**: Clean Tkinter interface dynamically adapts across screen resolutions and Windows display scaling (100%, 125%, 150%) without button clipping or font-rendering artifacts.
- **Privacy & Security First**: Your Gemini API key is stored locally on your machine in `.gemini_config.json` (gitignored). It is **never** included in exported datasets, Excel sheets, CSV, JSONL, HTML, or ZIP packages.

---

## 📐 Application Architecture & Workflow

```
┌────────────────────────────────────────────────────────────────────────┐
│                   Top Control & Dataset Management                     │
│  [+ New Dataset]  [Import Dataset]  [Export (ZIP)]  [Gemini API Key]  │
│  Source Name: [ biology ]    Next ID: [ biology_001 ]                  │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
       ┌────────────────────────────┴────────────────────────────┐
       ▼                                                         ▼
┌──────────────────────────────────────┐  ┌──────────────────────────────┐
│ Field 1: Question Input & Gemini     │  │ Field 2: Diagram / Visual    │
│ • Mode A: Screenshot OCR             │  │ • Paste/Browse visual figure │
│ • Mode B: Raw / Bijoy Text Furnisher │  │ • Embedded directly in Excel │
│ • Mode C: Both                       │  │ • Auto-saved to images/      │
│ • Live Bengali text editor (Ctrl+Z)  │  ├──────────────────────────────┤
└──────────────────────────────────────┘  │ Field 3: Answer Selection    │
                                          │ • Answer:                    │
                                          │   [ ক ]  [ খ ]  [ গ ]  [ ঘ ] │
                                          │ • Free-form custom answer    │
                                          ├──────────────────────────────┤
                                          │ [ Save & Add Question ]      │
                                          └──────────────────────────────┘
```

---

## 🛠️ Prerequisites & Installation

### Requirements
- **Python**: 3.10, 3.11, or 3.12
- **Operating System**: Windows 10 / 11 (also compatible with macOS and Linux)

### Setup Instructions

1. **Clone the repository**:
   ```bash
   git clone <your-repo-url>
   cd imageAdder
   ```

2. **(Recommended) Create and activate a virtual environment**:
   ```bash
   # Windows (PowerShell)
   python -m venv venv
   .\venv\Scripts\Activate.ps1

   # Windows (CMD)
   python -m venv venv
   .\venv\Scripts\activate.bat

   # Linux / macOS
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
   *Note: Only `pillow` and `openpyxl` are required. All Gemini API calls use Python's built-in `urllib` and `json` libraries with zero heavy external framework overhead.*

4. **Launch the Application**:
   ```bash
   python app.py
   ```

---

## 🔑 Gemini API Configuration

The application uses Google's latest Gemini models for text and multimodal OCR processing:

- **Default Primary Model**: `gemini-3.5-flash-lite` (Fast, cost-efficient, accurate)
- **Supported / Fallback Models**:
  - `gemini-flash-lite-latest`
  - `gemini-2.5-flash-lite`
  - `gemini-3.8-flash`
  - `gemini-flash-latest`

### How to Configure Your Key

1. Generate a free API key at [Google AI Studio](https://aistudio.google.com/app/apikey).
2. Click **Gemini API Key** in the top bar of the application.
3. Paste your key and select your preferred model, then click **Save Key**.
4. The key is securely saved locally in `.gemini_config.json` and remembered across sessions.

*(Alternative)* Set an environment variable before launching:
```bash
set GEMINI_API_KEY="your-api-key-here"
```

---

## 📖 Step-by-Step User Guide

### 1. Starting Up: Create or Import
When you launch the app, a minimal startup dialog offers two choices:
- **Create New Dataset (`Ctrl+N`)**: Enter a source name (e.g. `physics_paper1`). The app creates a clean timestamped workbook `data/physics_paper1_<timestamp>.xlsx` and initializes question IDs to `physics_paper1_001`.
- **Import Dataset (`Ctrl+O`)**: Select an existing `.xlsx` or `.csv` file. The app copies it into `data/` with a timestamp, copies any associated images, and loads the questions immediately.

### 2. Field 1: Question Input & Gemini OCR
Choose your input method using the mode buttons (`Screenshot`, `Raw / Bijoy Text`, or `Both`):
- **Screenshot Mode**:
  1. Capture a snip of the question using Snipping Tool / Lightshot (`Win + Shift + S`).
  2. Click the screenshot box or press `Ctrl+V`.
  3. With **Auto-extract on paste** checked (default), Gemini automatically transcribes the question into clean Unicode Bengali with structured options (`ক)`, `খ)`, `গ)`, `ঘ)`).
- **Raw / Bijoy Text Mode**:
  1. Paste legacy Bijoy/ANSI text into the box (`Ctrl+V`).
  2. Press `Ctrl+Enter` or click **Furnish Text with Gemini**.
  3. Gemini decodes the garbled text into modern Unicode Bengali.
- **Manual Adjustments**: The output text box supports full editing, undo (`Ctrl+Z`), and redo (`Ctrl+Y`).

### 3. Field 2: Diagram / Visual Image
If the question includes a diagram, chart, graph, or circuit:
1. Click the canvas or press **Paste** (`Ctrl+V` while focusing).
2. Alternatively, click **Browse** to select an image from disk.
3. To clear the visual, click **Clear**.

### 4. Field 3: Answer Selection
- Click **ক**, **খ**, **গ**, or **ঘ** for instant 1-click answer entry.
- Or type a custom numerical or text answer into the entry box.

### 5. Save & Continue (`Ctrl+Enter`)
- Press `Ctrl+Enter` or click **Save & Add Question to Dataset**.
- The app saves the entry to Excel, stores the image in `data/images/`, updates the status, advances the ID counter (`001` → `002`), and resets the input fields for the next question.

### 6. Exporting Your Dataset (`Ctrl+E`)
Click **Export (ZIP)** or press `Ctrl+E`:
- Automatically synchronizes and generates:
  - `<source>.xlsx` (with embedded image thumbnails)
  - `<source>.csv` (standard UTF-8 with BOM)
  - `<source>.jsonl` (standard AI benchmark format)
  - `<source>.html` (responsive visual preview with answer badges)
  - `images/` (all referenced visual diagram files)
- Packages all files into `data/exports/<source>_<timestamp>.zip`.
- Prompts to reveal the exported bundle in File Explorer.

---

## 📁 Repository & Directory Layout

```
├── app.py                      # Application launcher script
├── requirements.txt            # Minimal dependencies (Pillow, openpyxl)
├── README.md                   # Full documentation
├── .gitignore                  # Strict ignore rules for keys and dataset files
├── .gemini_config.json         # Local API key & model storage (gitignored)
│
├── core/                       # Core application logic
│   ├── models.py               # QuestionItem dataclass & filename sanitizers
│   ├── gemini_client.py        # Gemini REST client (zero external dependencies)
│   └── dataset_manager.py      # Excel, CSV, JSONL, HTML, and ZIP packaging
│
├── ui/                         # User Interface components
│   └── main_window.py          # Tkinter responsive UI & event handlers
│
├── tests/                      # Automated unit and integration tests
│   └── test_app2.py            # Comprehensive test suite (12 tests)
│
└── data/                       # Local workspace directory (gitignored data)
    ├── images/                 # Referenced question diagrams (.gitkeep tracked)
    └── exports/                # Exported ZIP archives (.gitkeep tracked)
```

---

## 📊 Dataset Output Specifications

### 1. Excel Workbook (`.xlsx`)
| Column | Header | Description |
| :---: | :--- | :--- |
| **A** | `id` | Unique question identifier (e.g. `biology_001`) |
| **B** | `question` | Full Bengali question text including options (Wrap Text enabled) |
| **C** | `answer` | Correct option letter (`ক`, `খ`, `গ`, `ঘ`) or written answer |
| **D** | `image` | Embedded visual thumbnail drawing (if present) |
| **E** | `image_path` | Relative path to image file (e.g. `images/biology_001.png`) |

### 2. JSONL Export (`.jsonl`)
Each line is a self-contained JSON object:
```json
{
  "id": "biology_001",
  "question": "কোষের শক্তিঘর কোনটি?\n\nক) প্লাস্টিড\nখ) মাইটোকন্ড্রিয়া\nগ) গলগি বস্তু\nঘ) রাইবোজোম",
  "answer": "খ",
  "image_path": "images/biology_001.png"
}
```

### 3. HTML Review Report (`.html`)
A standalone responsive webpage with embedded styles, Bangla font support, question cards, diagram thumbnails with modal previews, and answer key badges for rapid human verification.

---

## ⌨️ Keyboard Shortcuts Reference

| Shortcut | Context | Action |
| :--- | :--- | :--- |
| **Ctrl + N** | Global | Create a new timestamped dataset |
| **Ctrl + O** | Global | Import external `.xlsx` or `.csv` file |
| **Ctrl + E** | Global | Export dataset to `.zip` package |
| **Ctrl + S** | Global | Save & add question to dataset |
| **Ctrl + Enter** | Main Workspace | Save & add question to dataset |
| **Ctrl + Enter** | Raw Text Box | Furnish raw/Bijoy text with Gemini API |
| **Ctrl + V** | Question Box | Paste screenshot (auto-triggers extraction) |
| **Ctrl + Z** | Question Box | Undo text modification |
| **Ctrl + Y** | Question Box | Redo text modification |

---

## 🧪 Running Automated Tests

A comprehensive automated test suite validates data models, file sanitization, Excel/CSV/JSONL/HTML/ZIP generation, import workflows, and responsive UI scaling:

```bash
python -m unittest discover -s tests -p "test_*.py"
```

Expected output:
```text
Ran 12 tests in ~1.2s
OK
```

---

## 📄 License & Attribution

Designed and maintained for MMU Multimodal Dataset Creation projects.
Built with Python, Tkinter, Pillow, openpyxl, and Google Gemini API.
