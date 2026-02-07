#!/usr/bin/env python3

# PPTX to Images Converter

# Convert PowerPoint presentations to image files (PNG/JPG).
# Each slide is exported as a separate image file.

# Author: Avijit Roy
# GitHub: https://github.com/heyavijitroy
# Repository: https://github.com/heyavijitroy/pptx-to-pdf-converter
# License: MIT

# Requirements:
# - LibreOffice must be installed on your system
#   - Windows: Download from https://www.libreoffice.org/
#   - macOS: brew install --cask libreoffice
#   - Linux: sudo apt-get install libreoffice

# Usage:
#     python pptx_to_images.py
    
#     The script will convert all .pptx files in the current directory.
#     Images are saved in a folder named after each presentation.


import os
import subprocess
import sys
from pathlib import Path
import platform
import shutil

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


def convert_pptx_to_images(pptx_path, output_dir, soffice_path, image_format="png", resolution=150):
    """
    Convert a PowerPoint presentation to image files.
    
    Args:
        pptx_path: Path to the input .pptx file
        output_dir: Directory where images will be saved
        soffice_path: Path to the LibreOffice executable
        image_format: Output image format ('png' or 'jpg')
        resolution: DPI resolution (default: 150)
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    try:
        # Create output directory for this presentation
        pptx_name = pptx_path.stem
        image_dir = output_dir / pptx_name
        image_dir.mkdir(exist_ok=True)
        
        # First convert to PDF (intermediate step)
        temp_pdf = output_dir / f"{pptx_name}_temp.pdf"
        
        cmd_pdf = [
            soffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(pptx_path)
        ]
        
        result = subprocess.run(cmd_pdf, capture_output=True, text=True, timeout=120)
        
        if result.returncode != 0:
            print(f"  ✗ Failed to convert {pptx_path.name} to PDF")
            return False
        
        # Find the generated PDF
        pdf_file = output_dir / f"{pptx_name}.pdf"
        if not pdf_file.exists():
            print(f"  ✗ PDF file not created for {pptx_path.name}")
            return False
        
        # Convert PDF to images using LibreOffice's draw tool or pdftoppm
        # Since LibreOffice doesn't directly convert to images well, we'll use an alternative approach
        # We'll convert slides by importing PDF pages
        
        # Try using pdftoppm if available (Linux/Mac), otherwise use ImageMagick
        system = platform.system()
        
        # Method 1: Try pdftoppm (most reliable)
        try:
            # Check if pdftoppm is available
            check_cmd = ["which", "pdftoppm"] if system != "Windows" else ["where", "pdftoppm"]
            check_result = subprocess.run(check_cmd, capture_output=True, text=True)
            
            if check_result.returncode == 0:
                # Use pdftoppm to convert PDF to images
                pdftoppm_cmd = [
                    "pdftoppm",
                    "-" + image_format,
                    "-r", str(resolution),
                    str(pdf_file),
                    str(image_dir / "slide")
                ]
                result = subprocess.run(pdftoppm_cmd, capture_output=True, text=True, timeout=120)
                
                if result.returncode == 0:
                    # Clean up temporary PDF
                    pdf_file.unlink()
                    
                    # Count generated images
                    image_files = list(image_dir.glob(f"*.{image_format}"))
                    print(f"  ✓ Converted: {pptx_path.name} → {len(image_files)} images in '{pptx_name}/' folder")
                    return True
        except:
            pass
        
        # Method 2: Try ImageMagick's convert command
        try:
            check_cmd = ["which", "convert"] if system != "Windows" else ["where", "convert"]
            check_result = subprocess.run(check_cmd, capture_output=True, text=True)
            
            if check_result.returncode == 0:
                convert_cmd = [
                    "convert",
                    "-density", str(resolution),
                    str(pdf_file),
                    str(image_dir / f"slide.{image_format}")
                ]
                result = subprocess.run(convert_cmd, capture_output=True, text=True, timeout=120)
                
                if result.returncode == 0:
                    # Clean up temporary PDF
                    pdf_file.unlink()
                    
                    # Count generated images
                    image_files = list(image_dir.glob(f"*.{image_format}"))
                    print(f"  ✓ Converted: {pptx_path.name} → {len(image_files)} images in '{pptx_name}/' folder")
                    return True
        except:
            pass
        
        # If we reach here, keep the PDF and inform user
        print(f"  ⚠ PDF created but image conversion tools not found")
        print(f"    PDF saved as: {pdf_file.name}")
        print(f"    Install 'pdftoppm' or 'ImageMagick' for automatic image conversion")
        return True
        
    except subprocess.TimeoutExpired:
        print(f"  ✗ Timeout while converting {pptx_path.name}")
        return False
    except Exception as e:
        print(f"  ✗ Error converting {pptx_path.name}: {str(e)}")
        return False


def main():
    """Main function to convert PowerPoint presentations to images."""
    
    # Print header
    print(f"\n{'='*60}")
    print(f"PPTX to Images Converter v{__version__}")
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
    
    # Ask for image format preference
    print("Select output image format:")
    print("  1. PNG (higher quality, larger file size)")
    print("  2. JPG (smaller file size, good quality)")
    
    while True:
        format_choice = input("\nEnter your choice (1-2, default: 1): ").strip()
        if format_choice == "" or format_choice == "1":
            image_format = "png"
            break
        elif format_choice == "2":
            image_format = "jpg"
            break
        else:
            print("Invalid choice. Please enter 1 or 2.")
    
    # Ask for resolution
    print(f"\nEnter resolution (DPI):")
    print("  - 96: Screen quality (smaller files)")
    print("  - 150: Standard quality (recommended)")
    print("  - 300: High quality (larger files)")
    
    while True:
        res_input = input("\nEnter resolution (default: 150): ").strip()
        if res_input == "":
            resolution = 150
            break
        try:
            resolution = int(res_input)
            if 50 <= resolution <= 600:
                break
            else:
                print("Please enter a value between 50 and 600.")
        except ValueError:
            print("Invalid input. Please enter a number.")
    
    # Create output directory for images
    output_dir = script_dir / "slide_images"
    output_dir.mkdir(exist_ok=True)
    
    print(f"\n{'='*60}")
    print(f"Converting presentations to {image_format.upper()} images at {resolution} DPI...")
    print(f"{'='*60}\n")
    
    # Convert each file
    success_count = 0
    for pptx_file in pptx_files:
        if convert_pptx_to_images(pptx_file, output_dir, soffice_path, image_format, resolution):
            success_count += 1
    
    # Summary
    print(f"\n{'='*60}")
    print(f"Conversion complete: {success_count}/{len(pptx_files)} files converted successfully.")
    print(f"Images saved in: {output_dir.name}/ folder")
    print(f"{'='*60}")
    
    # Additional info about image conversion tools
    if success_count > 0:
        print("\nNote: For best results, consider installing:")
        print("  - pdftoppm (part of poppler-utils)")
        print("    Linux: sudo apt-get install poppler-utils")
        print("    macOS: brew install poppler")
        print("  - Or ImageMagick: https://imagemagick.org/")


if __name__ == "__main__":
    main()
