import PyInstaller.__main__
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    os.path.join(SCRIPT_DIR, "run_compliance.py"),
    "--onefile",
    "--name=Software Compliance Check",
    "--console",
    "--hidden-import=scan_software",
    "--hidden-import=check_upload",
    f"--add-data={os.path.join(SCRIPT_DIR, 'scan_software.py')};.",
    f"--add-data={os.path.join(SCRIPT_DIR, 'check_upload.py')};.",
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
    "--exclude-module=pygments",
    "--exclude-module=jinja2",
    "--exclude-module=tornado",
    "--exclude-module=zmq",
    "--exclude-module=cryptography",
    "--exclude-module=setuptools",
    "--clean",
])

print("\nBuild complete: dist/Software Compliance Check.exe")

print("  Users just double-click it — no Python needed.")
print("=" * 60)
