import PyInstaller.__main__
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

print("=" * 60)
print("  Building: Scan Software.exe")
print("=" * 60)

PyInstaller.__main__.run([
    os.path.join(SCRIPT_DIR, "scan_software.py"),
    "--onefile",
    "--name=Scan Software",
    "--console",
    f"--distpath={os.path.join(SCRIPT_DIR, 'dist')}",
    f"--workpath={os.path.join(SCRIPT_DIR, 'build')}",
    f"--specpath={SCRIPT_DIR}",
    "--exclude-module=torch",
    "--exclude-module=scipy",
    "--exclude-module=matplotlib",
    "--exclude-module=PIL",
    "--exclude-module=notebook",
    "--exclude-module=IPython",
    "--exclude-module=pytest",
    "--exclude-module=tkinter",
    "--exclude-module=sentence_transformers",
    "--exclude-module=httpx",
    "--exclude-module=fastapi",
    "--exclude-module=uvicorn",
    "--clean",
])

print("\nBuild complete: dist/Scan Software.exe")
print("  Copy 'dist/Scan Software.exe' for distribution to end users.")
print()
print("  NOTE: Users may see a Windows SmartScreen warning.")
print("  They should click 'More info' → 'Run anyway' to proceed.")
print("  Or right-click the exe → Properties → Unblock → Apply")
print("=" * 60)
