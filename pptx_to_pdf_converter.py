#!/usr/bin/env python3

# PPTX to PDF Converter

# Convert all PowerPoint (.pptx) files in the current directory to PDF format.
# This script uses LibreOffice in headless mode to preserve full presentation formatting
# including slides, images, layouts, and formatting.

# Author: Avijit Roy
# GitHub: https://github.com/heyavijitroy
# Repository: https://github.com/heyavijitroy/pptx-to-pdf-converter
# License: MIT

# Requirements:
# - LibreOffice must be installed on your system
#   - Windows: Download from https://www.libreoffice.org/
#   - macOS: brew install --cask libreoffice
#   - Linux: sudo apt-get install libreoffice (or equivalent)

# Usage:
#     python pptx_to_pdf_converter.py

# The script will automatically find and convert all .pptx files in the same directory.

import os
import subprocess
import sys
from pathlib import Path
import platform

__version__ = "1.0.0"
__author__ = "Avijit Roy"


def find_libreoffice():
    """Find the LibreOffice executable path based on the operating system."""
    system = platform.system()
    
    # Common LibreOffice paths for different operating systems
    possible_paths = []
    
    if system == "Windows":
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"C:\Program Files\LibreOffice 7\program\soffice.exe",
        ]
    elif system == "Darwin":  # macOS
        possible_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    else:  # Linux and others
        possible_paths = [
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
        ]
    
    # Check each possible path
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # Try to find it in PATH
    try:
        result = subprocess.run(
            ["which", "soffice"] if system != "Windows" else ["where", "soffice"],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            return result.stdout.strip().split('\n')[0]
    except:
        pass
    
    return None


def convert_pptx_to_pdf(pptx_path, output_dir, soffice_path):
    """
    Convert a PowerPoint presentation to PDF using LibreOffice.
    
    Args:
        pptx_path: Path to the input .pptx file
        output_dir: Directory where the PDF will be saved
        soffice_path: Path to the LibreOffice executable
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    try:
        # LibreOffice command to convert to PDF
        cmd = [
            soffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(pptx_path)
        ]
        
        # Run the conversion
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60  # 60 second timeout
        )
        
        if result.returncode == 0:
            pdf_name = pptx_path.stem + ".pdf"
            print(f"✓ Converted: {pptx_path.name} → {pdf_name}")
            return True
        else:
            print(f"✗ Failed to convert {pptx_path.name}")
            if result.stderr:
                print(f"  Error: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"✗ Timeout while converting {pptx_path.name}")
        return False
    except Exception as e:
        print(f"✗ Error converting {pptx_path.name}: {str(e)}")
        return False


def main():
    """Main function to convert all .pptx files in the current directory."""
    
    # Print header
    print(f"\n{'='*60}")
    print(f"PPTX to PDF Converter v{__version__}")
    print(f"Author: {__author__}")
    print(f"{'='*60}\n")
    
    # Find LibreOffice
    soffice_path = find_libreoffice()
    
    if not soffice_path:
        print("ERROR: LibreOffice not found!")
        print("\nPlease install LibreOffice:")
        print("  - Windows: https://www.libreoffice.org/download/download/")
        print("  - macOS: brew install --cask libreoffice")
        print("  - Linux: sudo apt-get install libreoffice")
        sys.exit(1)
    
    print(f"Using LibreOffice at: {soffice_path}\n")
    
    # Get the directory where the script is located
    script_dir = Path(__file__).parent.absolute()
    
    # Find all .pptx files in the directory
    pptx_files = list(script_dir.glob("*.pptx"))
    
    # Filter out temporary files (starting with ~$)
    pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]
    
    if not pptx_files:
        print("No .pptx files found in the current directory.")
        return
    
    print(f"Found {len(pptx_files)} PowerPoint file(s) to convert.\n")
    
    # Convert each file
    success_count = 0
    for pptx_file in pptx_files:
        if convert_pptx_to_pdf(pptx_file, script_dir, soffice_path):
            success_count += 1
    
    # Summary
    print(f"\n{'='*60}")
    print(f"Conversion complete: {success_count}/{len(pptx_files)} files converted successfully.")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
