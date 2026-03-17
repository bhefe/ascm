"""
Build 'Scan Software.exe' using a clean virtual environment to minimize exe size.
This avoids bundling the entire Anaconda environment (460+ MB → ~30 MB).
"""
import subprocess
import os
import sys
import shutil

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VENV_DIR = os.path.join(SCRIPT_DIR, ".build_venv")
DIST_DIR = os.path.join(SCRIPT_DIR, "dist")
BUILD_DIR = os.path.join(SCRIPT_DIR, "build")

# Only pdfplumber + openpyxl needed (pdfplumber pulls in pdfminer.six automatically)
REQUIRED_PACKAGES = ["pdfplumber", "openpyxl", "pyinstaller"]


def run_cmd(args, description=""):
    """Run a command and check for errors."""
    if description:
        print(f"  {description}...")
    result = subprocess.run(args, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ERROR: {result.stderr[:500]}")
        sys.exit(1)
    return result


def main():
    print("=" * 60)
    print("  Building: Scan Software.exe")
    print("=" * 60)

    # --- Step 1: Check PDFs ---
    approved_pdf = os.path.join(SCRIPT_DIR, "list of softwares (APPROVED by SAM).pdf")
    not_approved_pdf = os.path.join(SCRIPT_DIR, "LIST OF  SOFTWARES (NOT APPROVED).pdf")

    pdfs_found = []
    for pdf in [approved_pdf, not_approved_pdf]:
        if os.path.exists(pdf):
            pdfs_found.append(pdf)
            print(f"  [OK] Found: {os.path.basename(pdf)}")
        else:
            print(f"  [XX] Missing: {os.path.basename(pdf)}")

    # --- Step 2: Create clean venv ---
    print()
    venv_python = os.path.join(VENV_DIR, "Scripts", "python.exe")

    if not os.path.exists(venv_python):
        print("  [1/4] Creating clean virtual environment...")
        # Use base Python to create venv (not conda)
        base_python = sys.executable
        run_cmd([base_python, "-m", "venv", VENV_DIR])
        print(f"  [OK] venv created at: {VENV_DIR}")

        # --- Step 3: Install minimal deps ---
        print("  [2/4] Installing minimal dependencies...")
        venv_pip = os.path.join(VENV_DIR, "Scripts", "pip.exe")
        run_cmd([venv_pip, "install", "--no-cache-dir"] + REQUIRED_PACKAGES)
        print(f"  [OK] Installed: {', '.join(REQUIRED_PACKAGES)}")
    else:
        print("  [1/4] Using existing build venv")
        print("  [2/4] Dependencies already installed")

    # --- Step 4: Build with PyInstaller from venv ---
    print("  [3/4] Running PyInstaller...")

    venv_pyinstaller = os.path.join(VENV_DIR, "Scripts", "pyinstaller.exe")

    build_args = [
        venv_pyinstaller,
        os.path.join(SCRIPT_DIR, "scan_software.py"),
        "--onefile",
        "--name=Scan Software",
        "--console",
        f"--distpath={DIST_DIR}",
        f"--workpath={BUILD_DIR}",
        f"--specpath={SCRIPT_DIR}",
        "--hidden-import=pdfplumber",
        "--hidden-import=charset_normalizer",
        "--hidden-import=openpyxl",
        "--collect-all=charset_normalizer",
        "--collect-all=openpyxl",
        "--clean",
    ]

    for pdf in pdfs_found:
        build_args.append(f"--add-data={pdf}{os.pathsep}.")

    result = subprocess.run(build_args, text=True)
    if result.returncode != 0:
        print("  [XX] Build failed!")
        sys.exit(1)

    # --- Step 5: Report results ---
    exe_path = os.path.join(DIST_DIR, "Scan Software.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"\n  [4/4] Build successful!")
        print(f"  [OK] Output: {exe_path}")
        print(f"  [OK] Size: {size_mb:.1f} MB")
    else:
        print("  [XX] EXE not found after build!")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  Build complete: dist/Scan Software.exe")
    print("=" * 60)
    print("  Copy 'dist/Scan Software.exe' for distribution to end users.")
    print()
    print("  NOTE: Users may see a Windows SmartScreen warning.")
    print("  They should click 'More info' → 'Run anyway' to proceed.")
    print("  Or right-click the exe → Properties → Unblock → Apply")
    print("=" * 60)


if __name__ == "__main__":
    main()
