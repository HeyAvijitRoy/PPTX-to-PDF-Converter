# PPTX to PDF Converter

A simple Python script to batch convert PowerPoint presentations (.pptx) to PDF format while preserving all formatting, images, and layouts.

[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](https://github.com/heyavijitroy/pptx-to-pdf-converter)

## Features

- Batch convert all PPTX files in a directory
- Preserves original formatting, images, and layouts
- Cross-platform support (Windows, macOS, Linux)
- Uses only Python built-in libraries
- Simple command-line interface
- Automatic LibreOffice detection
- Progress feedback with conversion status

## Prerequisites

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
- No additional Python packages required! Uses only built-in libraries:
  - `os`
  - `subprocess`
  - `sys`
  - `pathlib`
  - `platform`

## Installation

1. Clone this repository:
```bash
git clone https://github.com/heyavijitroy/pptx-to-pdf-converter.git
cd pptx-to-pdf-converter
```

2. Make sure LibreOffice is installed (see Prerequisites above)

3. That's it! No additional dependencies to install.

## Usage

### Basic Usage

1. Place the `pptx_to_pdf_converter.py` script in the folder containing your .pptx files

2. Run the script:
```bash
python pptx_to_pdf_converter.py
```

3. The script will:
   - Find all .pptx files in the current directory
   - Convert each one to PDF format
   - Save the PDFs in the same directory
   - Display progress and results

### Example Output

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

## 📂 Project Structure

```
pptx-to-pdf-converter/
│
├── pptx_to_pdf_converter.py    # Main conversion script
├── README.md                    # This file
└── LICENSE                      # MIT License
```

## 🔧 How It Works

1. **LibreOffice Detection**: The script automatically detects LibreOffice installation on your system
2. **File Discovery**: Scans the current directory for all .pptx files
3. **Batch Conversion**: Uses LibreOffice's headless mode to convert each presentation
4. **Output**: Saves PDFs with the same filename in the same directory

The conversion preserves:
- All slides and their content
- Images and graphics
- Fonts and text formatting
- Layouts and designs
- Tables and charts
- Slide transitions and animations (in static form)

## Advanced Configuration

### Custom LibreOffice Path

If the script can't find your LibreOffice installation, you can modify the `find_libreoffice()` function to add your custom path.

### Conversion Timeout

The default timeout is 60 seconds per file. You can adjust this in the `convert_pptx_to_pdf()` function:

```python
timeout=60  # Change this value (in seconds)
```

## 🐛 Troubleshooting

### "LibreOffice not found" error

**Solution**: Make sure LibreOffice is installed and accessible from your system PATH. Try running:
```bash
# Linux/macOS
which soffice

# Windows (in Command Prompt)
where soffice
```

### Conversion fails for specific files

**Possible causes**:
- File is corrupted
- File is password-protected
- File is already open in PowerPoint
- Insufficient disk space

**Solution**: Close the file if it's open, check file integrity, and ensure you have write permissions.

### Script hangs or times out

**Solution**: Large presentations might take longer. Increase the timeout value in the script.

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

## 📚 Additional Resources

- [LibreOffice Documentation](https://documentation.libreoffice.org/)
- [Python subprocess module](https://docs.python.org/3/library/subprocess.html)
- [PowerPoint to PDF conversion best practices](https://www.libreoffice.org/)

## Changelog

### Version 1.0.0 (2025-02-07)
- Initial release
- Batch PPTX to PDF conversion
- Cross-platform support
- Automatic LibreOffice detection
- Progress feedback

---

**Note**: This script requires LibreOffice to be installed on your system. It does not require Microsoft Office or any paid software.
