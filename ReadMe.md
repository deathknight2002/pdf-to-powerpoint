# PDF to PowerPoint Converter

This script converts PDF files into PowerPoint presentations by converting each page to an image and embedding it in a slide. This ensures perfect visual fidelity without font or formatting issues, and the slides look exactly the same in PowerPoint as they did in PDF.

## Features

- ✅ **Perfect Visual Fidelity**: Converts PDF pages to images, eliminating font and formatting issues
- ✅ **High Quality Output**: Configurable DPI settings (default: 300 DPI)
- ✅ **Multiple Image Formats**: Support for PNG, JPEG, and TIFF formats
- ✅ **Progress Tracking**: Visual progress bars for large documents
- ✅ **Error Handling**: Comprehensive input validation and error messages
- ✅ **Flexible Output**: Custom output filenames and automatic overwrite protection
- ✅ **Performance Optimized**: Multi-threaded PDF processing
- ✅ **Professional CLI**: Full command-line interface with help and examples

## Why Use This Tool?

Many AV setups mean that if you are using PDF slides you cannot have 'next slide' on a confidence display whilst you speak. Having come up against this at multiple conferences, this tool was created to convert presentation decks from PDF to PowerPoint. Most converters try to keep the text as text but often introduce many inconsistencies. This tool ensures perfect visual reproduction.

## Installation

1. **Install Python Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Install System Dependencies**:
   - **macOS**: `brew install poppler`
   - **Ubuntu/Debian**: `sudo apt-get install poppler-utils`
   - **Windows**: Download and install poppler from [poppler for Windows](https://blog.alivate.com.au/poppler-windows/)

## Usage

### Basic Usage

```bash
# Convert PDF to PowerPoint (creates example.pptx)
python convert_improved.py example.pdf

# Original simple script (still available)
python convert.py example.pdf
```

### Advanced Usage

```bash
# Specify custom output filename
python convert_improved.py document.pdf -o presentation.pptx

# High resolution conversion (600 DPI)
python convert_improved.py document.pdf --dpi 600

# Use JPEG compression for smaller files
python convert_improved.py document.pdf --format JPEG

# Force overwrite existing files
python convert_improved.py document.pdf --force

# Verbose output with progress details
python convert_improved.py document.pdf --verbose

# Use more threads for faster processing
python convert_improved.py document.pdf --threads 4
```

### Command Line Options

```
positional arguments:
  pdf_file              Path to the PDF file to convert

options:
  -h, --help            Show help message and exit
  -o, --output OUTPUT   Output PowerPoint file path
  --dpi DPI             Resolution for conversion (default: 300)
  --format {PNG,JPEG,TIFF}  Image format for slides (default: PNG)
  --threads THREADS     Number of processing threads (default: 2)
  --force               Overwrite output file without confirmation
  --verbose             Enable verbose logging
  --version             Show program version
```

## Examples

```bash
# Basic conversion
python convert_improved.py presentation.pdf

# High-quality conversion with custom output
python convert_improved.py slides.pdf -o final_presentation.pptx --dpi 600

# Compressed output for email sharing
python convert_improved.py large_deck.pdf --format JPEG --force

# Batch processing with verbose output
python convert_improved.py *.pdf --verbose --threads 4
```

## File Formats

- **Input**: PDF files only
- **Output**: PowerPoint (.pptx) files
- **Image Formats**: 
  - PNG (default) - Best quality, larger files
  - JPEG - Good quality, smaller files
  - TIFF - Highest quality, largest files

## Performance Tips

- Use `--threads 4` or higher for large PDFs on multi-core systems
- Use `--format JPEG` for smaller output files when file size matters
- Use `--dpi 150` for draft conversions (faster processing)
- Use `--dpi 600` for high-quality printing

## Error Handling

The improved script includes comprehensive error handling:
- ✅ File existence validation
- ✅ PDF format verification
- ✅ Output path validation
- ✅ Overwrite protection
- ✅ Dependency checking
- ✅ Meaningful error messages

## Troubleshooting

### Common Issues

1. **"ModuleNotFoundError"**: Install requirements with `pip install -r requirements.txt`
2. **"poppler not found"**: Install poppler system dependency (see Installation section)
3. **"Permission denied"**: Check file permissions or use `--force` flag
4. **Large file processing**: Increase `--threads` or reduce `--dpi` for faster processing

### System Requirements

- Python 3.7+
- poppler-utils system package
- Sufficient disk space (output files can be 2-10x larger than input PDFs)

## Version History

- **v2.0** (Improved): Added CLI interface, error handling, progress bars, multiple formats
- **v1.0** (Original): Basic PDF to PowerPoint conversion functionality

## License

MIT License - Feel free to use and modify as needed.
