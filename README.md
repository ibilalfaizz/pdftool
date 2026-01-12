# File Format Converter

A fully functional Streamlit web application for converting files between different formats:
- **PowerPoint to PDF** (PPT/PPTX → PDF)
- **PDF to TIFF** (PDF → Multi-page TIFF)

## Features

### PowerPoint to PDF
- Upload `.ppt` or `.pptx` files
- Convert to PDF format
- Preview first page after conversion
- Download converted PDF

### PDF to TIFF
- Upload `.pdf` files
- Convert to multi-page TIFF format
- Adjustable DPI settings (150, 300, 600)
- Compression options (Deflate, LZW, Adobe Deflate, None)
- Preview first page after conversion
- Download converted TIFF

### UI/UX Features
- Clean, modern interface with gradient header
- Sidebar navigation between tools
- Progress indicators during conversion
- Success/error messages
- File size display (before and after conversion)
- Responsive design for desktop and mobile
- Session state management (remembers last DPI and compression settings)

## Installation

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Poppler (Required for PDF to TIFF conversion)

Poppler is required for the `pdf2image` library to work. Install it based on your operating system:

#### macOS
```bash
brew install poppler
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get update
sudo apt-get install poppler-utils
```

#### Linux (Fedora/CentOS/RHEL)
```bash
sudo yum install poppler-utils
```

#### Windows
1. Download Poppler for Windows from: https://github.com/oschwartz10612/poppler-windows/releases
2. Extract the zip file
3. Add the `bin` folder to your system PATH
   - Or set the environment variable: `POPPLER_PATH=C:\path\to\poppler\bin`

## Usage

### Running the Application

```bash
streamlit run app.py
```

The application will open in your default web browser at `http://localhost:8501`

### Using PowerPoint to PDF Converter

1. Select "PowerPoint to PDF" from the sidebar
2. Click "Choose a PowerPoint file" and upload a `.ppt` or `.pptx` file
3. Click "🔄 Convert to PDF"
4. Wait for the conversion to complete
5. Preview the first page (if available)
6. Click "📥 Download PDF" to save the converted file

### Using PDF to TIFF Converter

1. Select "PDF to TIFF" from the sidebar
2. Click "Choose a PDF file" and upload a `.pdf` file
3. Select your preferred DPI (150, 300, or 600)
   - **150 DPI**: Lower quality, smaller file size
   - **300 DPI**: Recommended for most uses
   - **600 DPI**: High quality, larger file size
4. Choose compression method:
   - **tiff_deflate**: Recommended, good balance
   - **tiff_lzw**: Alternative compression
   - **tiff_adobe_deflate**: Adobe-specific compression
   - **None**: No compression, largest file size
5. Click "🔄 Convert to TIFF"
6. Wait for the conversion to complete
7. Preview the first page
8. Click "📥 Download TIFF" to save the converted file

## Technical Details

### Dependencies

- **streamlit**: Web framework for the application
- **python-pptx**: Reading and parsing PowerPoint files
- **reportlab**: PDF generation
- **pdf2image**: Converting PDF pages to images
- **Pillow (PIL)**: Image processing and TIFF creation

### File Processing

- Files are processed in memory using temporary buffers
- Temporary files are cleaned up after processing
- Multi-page PDFs and TIFFs are fully supported
- Large files are handled efficiently

### Limitations

- **PowerPoint to PDF**: Currently extracts text content and creates a formatted PDF. For full visual rendering (images, complex layouts), consider using LibreOffice headless mode or cloud conversion services.
- **PDF to TIFF**: Requires Poppler to be installed on the system.

## Troubleshooting

### "Required libraries not installed" error
- Make sure you've installed all dependencies: `pip install -r requirements.txt`

### PDF to TIFF conversion fails
- Verify Poppler is installed: `pdftoppm -h` (should show help)
- On Windows, ensure Poppler is in your PATH or set `POPPLER_PATH` environment variable
- Check that the PDF file is not corrupted or password-protected

### PowerPoint conversion shows only text
- This is expected behavior with the current implementation
- The converter extracts text content and formats it in PDF
- For full visual conversion, consider using LibreOffice or other tools

### Memory errors with large files
- The app processes files in memory
- For very large files (>100MB), consider splitting them first
- Ensure you have sufficient RAM available

## License

This project is provided as-is for personal and educational use.

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

