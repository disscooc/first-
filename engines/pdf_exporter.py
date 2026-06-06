import os, subprocess, sys, platform
from pathlib import Path

def find_libreoffice():
    system = platform.system()
    if system == "Windows":
        paths = []
        for base in [os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)")]:
            if base:
                paths.append(str(Path(base) / "LibreOffice" / "program" / "soffice.exe"))
        for p in paths:
            if os.path.exists(p):
                return p
        # Check PATH
        try:
            result = subprocess.run(["where", "soffice"], capture_output=True, text=True, shell=True)
            if result.stdout.strip():
                return result.stdout.strip().split("\n")[0]
        except:
            pass
    elif system == "Darwin":
        p = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(p):
            return p
        try:
            result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
            if result.stdout.strip():
                return result.stdout.strip()
        except:
            pass
    else:  # Linux
        try:
            result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
            if result.stdout.strip():
                return result.stdout.strip()
        except:
            pass
    return None

def convert_to_pdf(input_path, output_dir=None):
    lo = find_libreoffice()
    if not lo:
        raise RuntimeError("LibreOffice not found. Please install LibreOffice for PDF export.\n"
                           "Download: https://www.libreoffice.org/download/")
    if output_dir is None:
        output_dir = os.path.dirname(input_path)
    os.makedirs(output_dir, exist_ok=True)
    args = [lo, "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_path]
    try:
        result = subprocess.run(args, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
        base = os.path.splitext(os.path.basename(input_path))[0]
        pdf_path = os.path.join(output_dir, base + ".pdf")
        if os.path.exists(pdf_path):
            return pdf_path
        raise RuntimeError(f"PDF not generated at {pdf_path}")
    except subprocess.TimeoutExpired:
        raise RuntimeError("PDF conversion timed out (120s). File may be too large.")

def is_available():
    return find_libreoffice() is not None

