# Khmer-English OCR Application

A modern OCR (Optical Character Recognition) application that supports Khmer and English text extraction from images and PDF files. Built with Python, Tesseract (pytesseract), and a clean, minimal UI. Optional AI proofreading (OpenAI) can automatically correct OCR mistakes in Khmer + English.

## Features

- **Multi-language OCR**: Khmer and English via Tesseract (can add more with traineddata)
- **Multiple file formats**: Process images (PNG, JPG, JPEG, BMP, TIFF, WEBP) and PDF files
- **Lightweight OCR engine**: Uses system Tesseract via pytesseract
- **Real-time processing stats**: File size, pages processed, characters extracted, processing speed
- **Modern UI**: Clean, minimal design with progress indicators and status updates
- **Export options**: Save extracted text as TXT or DOCX files
- **Automatic language detection**: Intelligently detects mixed language content
 - **AI proofreading (optional)**: One-click “AI Proofread” button to fix OCR mistakes (requires OpenAI API key)
  

## Requirements

- Python 3.8+
- Tesseract OCR binary installed on your system
- Poppler (for PDF processing)
 - Optional: OpenAI account + API key for AI proofreading

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **(Optional) Enable AI proofreading:**

   - Create an OpenAI API key: https://platform.openai.com/
   - Set the environment variable before running the app (replace with your key):
     - macOS/Linux (bash/zsh):
       ```bash
       export OPENAI_API_KEY="sk-..."
       ```
     - Windows (PowerShell):
       ```powershell
       setx OPENAI_API_KEY "sk-..."
       ```
   - The app uses model `gpt-4o-mini` by default. You can change it in `app.py`.

2. **Install Poppler (for PDF support):**
   
   **macOS:**
   ```bash
   brew install poppler
   ```
   
   **Ubuntu/Debian:**
   ```bash
   sudo apt-get install poppler-utils
   ```

3. **Install Tesseract:**

   - macOS (Homebrew):
     ```bash
     brew install tesseract
     ```
   - Ubuntu/Debian:
     ```bash
     sudo apt-get update && sudo apt-get install -y tesseract-ocr
     ```
   - Windows: Install from https://github.com/UB-Mannheim/tesseract/wiki

4. **Install Khmer language data (khm):**

   Tesseract needs `khm.traineddata` for Khmer. If not already present:

   - Download from:
     - Best accuracy: https://github.com/tesseract-ocr/tessdata_best
     - Balanced: https://github.com/tesseract-ocr/tessdata
   - Place `khm.traineddata` into your tessdata folder. Common paths:
     - macOS ARM (Homebrew): `/opt/homebrew/share/tessdata`
     - macOS Intel (Homebrew): `/usr/local/share/tessdata`
     - Linux: `/usr/share/tesseract-ocr/4.00/tessdata` or `/usr/share/tesseract-ocr/tessdata`
     - Windows: `C:\\Program Files\\Tesseract-OCR\\tessdata`
   
   Verify installation:
   ```bash
   tesseract --list-langs | grep -i khm
   ```

## Usage

Run the application:
```bash
python app.py
```

1. Click "Choose File" to select an image or PDF
2. The application will automatically detect the language
3. View real-time processing statistics
4. Save the extracted text as TXT or DOCX file
 5. (Optional) Click "AI Proofread" to let AI correct OCR mistakes in the extracted text

## Technical Details

- **OCR Engine**: Tesseract via pytesseract; uses `eng` and `khm` (or `khm+eng` for mixed)
- **Language Detection**: Simple heuristic; you can manually choose by editing code
- **Image Processing**: Lightweight preprocessing to improve OCR readability
- **UI Framework**: ttkbootstrap for modern, responsive interface
 - **AI Proofreading**: Implemented via OpenAI Python SDK. Reads `OPENAI_API_KEY` from environment and calls `gpt-4o-mini` with low temperature for conservative corrections.
# Arn-AI

## Author

- Developer: PHAL PHEAKDEY

