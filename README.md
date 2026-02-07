# PowerPoint Automation Toolkit

A collection of Python scripts for automating PowerPoint presentation tasks including format conversion, merging, and image extraction.

[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](https://github.com/heyavijitroy/pptx-to-pdf-converter)

## Features

### PPTX to PDF Converter
- Batch convert all PPTX files to PDF format
- Preserves original formatting, images, and layouts
- Cross-platform support (Windows, macOS, Linux)
- Zero dependencies - uses only Python built-in libraries

### PPTX Merger
- Combine multiple presentations into one file
- Choose merge order (alphabetical or custom)
- Optional separator slides between presentations
- Interactive command-line interface

### PPTX to Images Converter
- Export each slide as PNG or JPG images
- Customizable resolution (DPI settings)
- Organized output in separate folders per presentation
- Perfect for creating thumbnails or social media content

## Prerequisites

### System Requirements

**For PPTX to PDF & PPTX to Images:**
- LibreOffice (free and open-source)
  
**For PPTX Merger:**
- Python library: `python-pptx`

**For enhanced image conversion (optional):**
- `pdftoppm` (from poppler-utils) or ImageMagick

### LibreOffice Installation

This script uses LibreOffice in headless mode to perform the conversions. You need to install LibreOffice on your system:

#### Windows
Download and install from [LibreOffice official website](https://www.libreoffice.org/download/download/)

#### macOS
```bash
brew install --cask libreoffice
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get update
sudo apt-get install libreoffice
```

#### Linux (Fedora/RHEL)
```bash
sudo dnf install libreoffice
```

### Python Requirements

- Python 3.6 or higher

**Built-in libraries (no installation needed):**
- `os`, `subprocess`, `sys`, `pathlib`, `platform`

**External libraries (install as needed):**
```bash
# For PPTX Merger only
pip install python-pptx

# For enhanced image conversion (optional)
# Linux/Ubuntu
sudo apt-get install poppler-utils

# macOS
brew install poppler
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/heyavijitroy/pptx-to-pdf-converter.git
cd pptx-to-pdf-converter
```

2. Install system dependencies:
```bash
# Install LibreOffice (see Prerequisites above)

# Install Python dependencies
pip install -r requirements.txt
```

3. You're ready to go!

## Usage

### 1. PPTX to PDF Converter

Convert all PowerPoint files to PDF format:

```bash
python pptx_to_pdf_converter.py
```

**What it does:**
- Finds all .pptx files in the current directory
- Converts each to PDF format
- Saves PDFs in the same directory

**Example Output:**
```
============================================================
PPTX to PDF Converter v1.0.0
Author: Avijit Roy
============================================================

Using LibreOffice at: /usr/bin/soffice

Found 3 PowerPoint file(s) to convert.

✓ Converted: presentation1.pptx → presentation1.pdf
✓ Converted: meeting_notes.pptx → meeting_notes.pdf
✓ Converted: project_proposal.pptx → project_proposal.pdf

============================================================
Conversion complete: 3/3 files converted successfully.
============================================================
```

---

### 2. PPTX Merger

Combine multiple PowerPoint files into one:

```bash
python merge_pptx.py
```

**Interactive Options:**
1. Merge all files in alphabetical order
2. Merge files in custom order
3. Select specific files to merge
4. Add separator slides between presentations (optional)

**Example Session:**
```
============================================================
PPTX Merger v1.0.0
Author: Avijit Roy
============================================================

Found 3 PowerPoint file(s) in the current directory.

Available PowerPoint files:
  1. intro.pptx
  2. main_content.pptx
  3. conclusion.pptx

Options:
  1. Merge all files in alphabetical order
  2. Merge all files in custom order
  3. Select specific files to merge

Enter your choice (1-3): 2
Enter file numbers in desired order (space-separated): 1 2 3

Add separator slides between presentations? (y/n): y

Enter output filename (default: merged_presentation.pptx): final_deck.pptx

============================================================
Merging 3 presentations...
============================================================

  Base: intro.pptx (5 slides)
  Added: main_content.pptx (15 slides)
  Added: conclusion.pptx (3 slides)

✓ Successfully merged into: final_deck.pptx
  Total slides: 26

============================================================
Merge completed successfully! 🎉
============================================================
```

---

### 3. PPTX to Images Converter

Export each slide as an image file:

```bash
python pptx_to_images.py
```

**Interactive Options:**
1. Choose image format (PNG or JPG)
2. Set resolution/quality (DPI: 96-600)
3. Automatic folder organization

**Example Session:**
```
============================================================
PPTX to Images Converter v1.0.0
Author: Avijit Roy
============================================================

Using LibreOffice at: /usr/bin/soffice

Found 2 PowerPoint file(s) to convert.

Select output image format:
  1. PNG (higher quality, larger file size)
  2. JPG (smaller file size, good quality)

Enter your choice (1-2, default: 1): 1

Enter resolution (DPI):
  - 96: Screen quality (smaller files)
  - 150: Standard quality (recommended)
  - 300: High quality (larger files)

Enter resolution (default: 150): 150

============================================================
Converting presentations to PNG images at 150 DPI...
============================================================

  ✓ Converted: product_demo.pptx → 12 images in 'product_demo/' folder
  ✓ Converted: training.pptx → 8 images in 'training/' folder

============================================================
Conversion complete: 2/2 files converted successfully.
Images saved in: slide_images/ folder
============================================================
```

**Output Structure:**
```
slide_images/
├── product_demo/
│   ├── slide-1.png
│   ├── slide-2.png
│   └── ...
└── training/
    ├── slide-1.png
    ├── slide-2.png
    └── ...
```

## Project Structure

```
pptx-automation-toolkit/
│
├── pptx_to_pdf_converter.py    # Convert PPTX to PDF
├── merge_pptx.py                # Merge multiple PPTX files
├── pptx_to_images.py            # Export slides as images
├── requirements.txt             # Python dependencies
├── README.md                    # This file
├── LICENSE                      # MIT License
└── .gitignore                   # Git ignore rules
```

## How It Works

### PPTX to PDF Converter
1. **LibreOffice Detection**: Automatically finds LibreOffice on your system
2. **File Discovery**: Scans directory for .pptx files
3. **Batch Conversion**: Uses LibreOffice headless mode
4. **Output**: Saves PDFs with identical filenames

### PPTX Merger  
1. **File Selection**: Lists all available presentations
2. **User Choice**: Interactive menu for merge options
3. **Slide Copying**: Copies all slides maintaining formatting
4. **Optional Separators**: Adds labeled divider slides
5. **Output**: Single merged presentation file

### PPTX to Images
1. **PDF Intermediate**: Converts PPTX to PDF first
2. **Image Extraction**: Uses pdftoppm or ImageMagick
3. **Organization**: Creates folders per presentation
4. **Quality Control**: Customizable DPI settings

**What gets preserved in conversions:**
- All slides and their content
- Images and graphics
- Fonts and text formatting
- Layouts and designs
- Tables and charts
- Slide transitions (in static form)

## Advanced Configuration

### Custom LibreOffice Path

If the script can't find your LibreOffice installation, you can modify the `find_libreoffice()` function to add your custom path.

### Conversion Timeout

The default timeout is 60 seconds per file. You can adjust this in the `convert_pptx_to_pdf()` function:

```python
timeout=60  # Change this value (in seconds)
```

## Troubleshooting

### Common Issues

#### "LibreOffice not found" error

**Solution**: Make sure LibreOffice is installed and accessible from your system PATH.
```bash
# Linux/macOS - Check if installed
which soffice

# Windows (in Command Prompt)
where soffice
```

#### "python-pptx not installed" (for Merger)

**Solution**: Install the required library
```bash
pip install python-pptx
```

#### Images not generating (PPTX to Images)

**Solution**: Install additional tools for better image conversion
```bash
# Linux/Ubuntu
sudo apt-get install poppler-utils

# macOS  
brew install poppler

# Or install ImageMagick
# Linux
sudo apt-get install imagemagick
# macOS
brew install imagemagick
```

#### Conversion fails for specific files

**Possible causes**:
- File is corrupted
- File is password-protected
- File is currently open in PowerPoint
- Insufficient disk space

**Solution**: 
- Close the file if it's open
- Check file integrity
- Ensure write permissions
- Free up disk space

#### Script hangs or times out

**Solution**: Large presentations take longer. Increase timeout in the script:
```python
timeout=120  # Increase from 60 to 120 seconds
```

#### Merged presentation has formatting issues

**Solution**: 
- Ensure all source files use compatible themes
- Check that all fonts are installed on your system
- Try merging without separator slides first

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/heyavijitroy/pptx-to-pdf-converter/issues).

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

**Avijit Roy**

- GitHub: [@heyavijitroy](https://github.com/heyavijitroy)

## Show your support

Give a ⭐️ if this project helped you!

## Additional Resources

- [LibreOffice Documentation](https://documentation.libreoffice.org/)
- [Python subprocess module](https://docs.python.org/3/library/subprocess.html)
- [PowerPoint to PDF conversion best practices](https://www.libreoffice.org/)

## Changelog

### Version 1.1.0 (2025-02-07)
- PPTX to PDF Converter - Initial release
- PPTX Merger - Combine multiple presentations
- PPTX to Images - Export slides as PNG/JPG
- Cross-platform support (Windows, macOS, Linux)
- Interactive command-line interfaces
- Comprehensive documentation

---

**Note**: These scripts use free and open-source tools. No Microsoft Office license required!