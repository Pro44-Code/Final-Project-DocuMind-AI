# DocuMind AI

DocuMind AI is a privacy-first, fully localized Document Management System (DMS) designed to streamline document workflows for smaller enterprises and teams. Developed as part of an academic thesis project, the system integrates text extraction, optical character recognition (OCR), hybrid semantic search, and intelligent text summarization into a single desktop application—operating entirely offline to ensure complete data sovereignty.


## 🚀 Key Features

- **Privacy-First Architecture:** Powered by **Ollama**, the large language model (LLM) runs 100% locally on your machine. Confidential legal, financial, or HR documents are never transmitted over the internet.
- **Two-Level Performance Caching:** 
  - *Level 1:* Checks file modification timestamps to prevent redundant directory re-scanning.
  - *Level 2:* Uses **MD5 hashing** on document text. If a file's hash matches an existing log entry, pre-generated AI summaries are loaded instantly, dramatically reducing CPU load.
- **Robust OCR Fallback:** Integrates **Tesseract OCR** to parse text from scanned or low-resolution image-based PDF files with high precision.
- **Hybrid Text Retrieval:** Combines statistical and keyword-matching mechanisms (**TF-IDF matrix**) for lightning-fast, content-based document searching.
- **Automated Deadline Tracking:** Parsed via `openpyxl`, the system automatically detects, isolates, and monitors critical expiration dates and project milestones embedded in spreadsheets.

### Technical Stack

- `PyQt5` for desktop UI
- `PyMuPDF (fitz)` for PDF text extraction
- `pytesseract` + `Pillow` + Tesseract OCR for scanned PDFs
- `python-docx` for DOCX reading/writing
- `openpyxl` for XLSX processing
- `Ollama` local model runtime for AI tasks
- `requests.Session` for model API communication
- JSON-based caching in `cache` folder

### Architecture Notes

- **Privacy-first local execution**
- **Cache-aware scanning and AI response reuse**
- **Heuristic OCR triggering** (OCR only when native PDF text looks insufficient)
- **Parallel file parsing** for faster scanning on multi-core CPUs
- **AI history tracking** for summarize/generate operations, viewable directly in the app UI
- **Threaded background AI requests to keep UI responsive**

### Run Instructions

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

OCR prerequisites (Windows):

1. Install Tesseract OCR: [UB Mannheim build](https://github.com/UB-Mannheim/tesseract/wiki)  
2. Add Georgian data file: [kat.traineddata](https://github.com/tesseract-ocr/tessdata/raw/main/kat.traineddata) into `C:\Program Files\Tesseract-OCR\tessdata`  
3. Verify:

```powershell
tesseract --list-langs
```

### Detailed Setup Notes

1. Ensure Python 3.10+ is installed and available in `PATH`.
2. Create and activate a virtual environment before installing dependencies.
3. Install and run Ollama locally for AI features.
4. Pull at least one model (default is `llama3.2:1b` unless overridden by env vars).
5. Launch the app and validate search on a sample folder.

(Optional) Test file generation:

```bash
python file_generator/files.py
```

### Future Work (Roadmap)

- **Distributed deployment:** server mode for shared network folders and team usage.
- **Asynchronous indexing:** true background scanning with queue management and progress tracking.
- **Semantic retrieval:** embedding-based ranking and hybrid retrieval pipeline.
- **Advanced multilingual support:** improved Georgian handling and multi-language prompt strategies.
- **Document understanding:** clause detection and risk scoring.
- **Evaluation framework:** benchmark datasets, quality metrics, and regression testing for model outputs.