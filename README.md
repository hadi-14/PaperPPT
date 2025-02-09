# PaperPPT

A Python application that converts CAIE Multiple Choice Question (MCQ) PDF papers into PowerPoint presentations, with each question on a separate slide. Perfect for teachers and educators who want to review MCQ papers with their students.

## Features

- Convert MCQ PDF papers to PowerPoint presentations
- Support for both single file and batch processing
- Automatic question detection and extraction
- Preserves images and formatting from the original PDF
- Customizable slide timing for automated presentations
- User-friendly GUI interface
- Support for command-line usage

## Installation

Download the latest release `MCQs_to_PPT.exe` from the releases page. No installation required - the application is portable and can be run directly.

If you want to run from source or build yourself:

1. Clone the repository:

```bash
git clone https://github.com/yourusername/PaperPPT.git
cd PaperPPT
```

2. Install required dependencies:

```bash
pip install -r requirements.txt
```

3. Optional: Build the executable using Nuitka:

```bash
python -m nuitka MCQs_to_PPT.py --onefile --enable-plugin=tk-inter --windows-console-mode=disable --include-data-dir=templates=templates
```

## Usage

### GUI Mode

1. Launch `MCQs_to_PPT.exe`
2. Choose between single file or batch processing mode
3. Select input PDF file(s) and output location(s)
4. Optional: Set slide timing duration (in seconds)
5. Click "Start Processing"

### Command Line Mode

No CLI mode in exe form:

```bash
python MCQQuestionSplitter.py [pdf_path] [--output OUTPUT] [--seconds SECONDS]
```

Arguments:

- `pdf_path`: Path to the PDF file
- `--output`, `-o`: Output PowerPoint file name (default: mcq_presentation.pptx)
- `--seconds`, `-s`: Number of seconds each slide should display (default: None for manual control)

## Project Structure

```
PaperPPT/
├── papers/              # Input PDF files
├── ppts/               # Output PowerPoint files
├── templates/          # PowerPoint templates
├── MCQs_to_PPT.py     # Main application source
├── MCQQuestionSplitter.py  # Core conversion logic
└── MCQs_to_PPT.exe    # Compiled executable
```

## Requirements

- Windows 7 or later
- No additional software required when using the executable
- For source code:
  - Python 3.8+
  - pdfplumber
  - python-pptx
  - Pillow
  - tqdm
  - customtkinter

## Building from Source

The executable is compiled using Nuitka with the following command:

```bash
python -m nuitka MCQs_to_PPT.py --onefile --enable-plugin=tk-inter --windows-console-mode=disable --include-data-dir=templates=templates
```

This creates a single executable file that includes all necessary dependencies.

## Limitations

- Currently optimized for standard MCQ format with questions numbered 1-40
- PDF must have clear question numbering and formatting
- Best results with single-column layout PDFs
- Images must be embedded in the PDF (not linked)

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- PDF processing powered by [pdfplumber](https://github.com/jsvine/pdfplumber)
- PowerPoint generation using [python-pptx](https://python-pptx.readthedocs.io/)
