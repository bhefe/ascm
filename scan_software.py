import winreg
import csv
import os
import sys
import socket
import getpass
import subprocess
import json
import re
import io
from datetime import datetime

import pdfplumber


# When running as frozen exe, PDFs are bundled in _MEIPASS temp dir
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NOT_APPROVED_PDFS = [
    os.path.join(BASE_DIR, "LIST OF  SOFTWARES (NOT APPROVED).pdf")
]

# ── Pre-installed / system software to ignore (won't appear in report) ──
IGNORE_LIST = [
    "241916F58D6E7",
    "7.1 Surround Sound",
    "Actions Server",
    "Adobe Acrobat DC",
    "Adobe After Effects",
    "Adobe Genuine Service",
    "Adobe Media Encoder",
    "Adobe Photoshop",
    "Adobe Premiere Pro",
    "AI Noise Cancelation",
    "aimgr",
    "Antigravity (User)",
    "Application Compatibility Enhancements",
    "Armoury Crate",
    "ASUS Aac",
    "ASUS Ambient",
    "ASUS AURA",
    "ASUS Device Check",
    "ASUS Framework",
    "ASUS Hotplug",
    "ASUS Keyboard",
    "ASUS MB",
    "ASUS Mouse",
    "ASUS Smart",
    "ASUS Update",
    "ASUSAmbient",
    "ASUSPCAssistant",
    "AURA",
    "AVCEncoder Video Extension",
    "Bing Search",
    "Canon IJ Scan",
    "Cisco AnyConnect",
    "Clipchamp",
    "Composer - PHP",
    "Copilot",
    "Configuration Manager",
    "Cross Device",
    "Dev-C++",
    "DiagnosticsHub",
    "Dolby Vision",
    "DTSXUltra",
    "EA app",
    "EA SPORTS",
    "ELANTrack",
    "Entity Framework",
    "Epic Games",
    "EUC Wallpaper",
    "Game Assist",
    "Gaming App",
    "Gaming Services",
    "GameSDK",
    "Glanceby",
    "Glidex",
    "GlideX",
    "HPPrinter",
    "i Cloud",
    "InputMapper",
    "Intel Arc",
    "Intel Graphics",
    "Invgate Asset",
    "InvGate Insight",
    "It Takes Two",
    "Ivanti Secure",
    "Kinect for Windows",
    "Language Experience",
    "Launcher Prerequisites",
    "Lightshot",
    "Lenovo Companion",
    "Lenovo Migration",
    "Lenovo Quick",
    "Lenovo Settings",
    "Lenovo System Update",
    "Lenovo Vantage",
    "M365Companions",
    "Messaging",
    "Microsoft .NET",
    "Microsoft 365 Apps",
    "Microsoft ASP.NET",
    "Microsoft Azure",
    "Microsoft Edge",
    "Microsoft Family",
    "Microsoft GameInput",
    "Microsoft Intune",
    "Microsoft Power BI",
    "Microsoft ODBC",
    "Microsoft OneNote",
    "Microsoft SQL Server Compact",
    "Microsoft Office Hub",
    "Microsoft OneDrive",
    "Microsoft Policy",
    "Microsoft Server Speech",
    "Microsoft Teams",
    "Microsoft Visual Studio",
    "MSTeams",
    "Mozilla Maintenance",
    "NVIDIA",
    "Office 16 Click-to-Run",
    "Office Push Notification",
    "One Connect",
    "One Note",
    "One Note Virtual",
    "OTA Dependencies",
    "Outlook For Windows",
    "Overcooked",
    "Parsec",
    "PersonalGatewayComponents",
    "POV-Ray",
    "Power Automate",
    "Portal",
    "Print3D",
    "PuTTY",
    "Pulse Application",
    "Pulse Secure",
    "Quick Assist",
    "Realtek Audio",
    "RefreshRateService",
    "Rockstar Games",
    "ROG Live Service",
    "SAP Business",
    "Screen Pad Master",
    "SketchUp",
    "Steam",
    "Sec Health UI",
    "Shell Extension",
    "SmartPSS",
    "Start Experiences",
    "Synaptics Control",
    "Synaptics Utilities",
    "Teams Machine-Wide",
    "Tesseract-OCR",
    "Thunderbolt Control",
    "Tools for .Net",
    "TypeScript SDK",
    "Ubisoft Connect",
    "Universal CRT Tools",
    "VcXsrv",
    "Visual Studio Build Tools",
    "Visual Studio Community",
    "vs_clickoncesigntoolmsi",
    "vs_communitysharedmsi",
    "vs_filehandler_x86",
    "vs_minshellinteropmsi",
    "vs_minshellmsi",
    "vs_minshellmsires",
    "vs_minshellsharedmsi",
    "Web Experience",
    "Whiteboard",
    "WebStorm",
    "Widgets Platform",
    "Windows HDRCalibration",
    "Windows Subsystem For Linux",
    "WinRT Intellisense",
    "Wireshark",
    "WoJ XInput Emulator",
    "X-Rite Color",
    "Stable",
    "Update for",
    "Google Chrome",
    "Adobe Acrobat Reader Core App",
]


def get_installed_software():
    software_list = []
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]

    for hive, path in reg_paths:
        try:
            key = winreg.OpenKey(hive, path)
        except OSError:
            continue

        for i in range(winreg.QueryInfoKey(key)[0]):
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)

                try:
                    name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                except OSError:
                    name = None

                try:
                    version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                except OSError:
                    version = ""

                try:
                    publisher = winreg.QueryValueEx(subkey, "Publisher")[0]
                except OSError:
                    publisher = ""

                try:
                    install_date = winreg.QueryValueEx(subkey, "InstallDate")[0]
                except OSError:
                    install_date = ""

                if name and not _should_exclude(name, version, publisher):
                    software_list.append({
                        "Software": name.strip(),
                        "Version": str(version).strip(),
                        "Publisher": str(publisher).strip(),
                        "InstallDate": str(install_date).strip(),
                    })

                subkey.Close()
            except OSError:
                continue

        key.Close()

    return software_list


def _should_exclude(name: str, version: str, publisher: str) -> bool:
    """Filter out system components, drivers, and runtimes that SAM doesn't care about."""
    name_lower = name.lower()
    version_str = str(version).lower()
    publisher_lower = str(publisher).lower() if publisher else ""

    # Skip corrupted/invalid entries
    if name_lower in ("nan", "none", "") or name.startswith(("(", "{", "!")):
        return True

    # Skip entries that are just GUIDs
    if re.match(r'^[a-f0-9\-]{36}$|^[a-f0-9]{32}$', name_lower.strip()):
        return True

    # Skip Windows system components
    if any(x in name_lower for x in [
        "windows ", "win32 ", "system32", "drivers", "device", "helper",
        "sec health", "web experience", "widgets platform", "pc health check",
        "language experience pack", "keyboard layout", "ime",
        "update for microsoft", "microsoft update health",
    ]):
        return True

    # Skip drivers
    if any(x in name_lower for x in [
        " driver", " drivers", " mp drivers", " graphics", " chipset",
        "realtek", "nvidia", "intel graphics", "amd", "broadcom",
    ]):
        return True

    # Skip redistributables and runtimes
    if any(x in name_lower for x in [
        "visual c++", "vcredist", ".net runtime", ".net framework",
        "dotnet", "msvc", "redistributable", "runtime", "webrtc",
        "directx", "vulkan", "opengl", "dx12",
    ]):
        return True

    # Skip Office internal components
    if any(x in name_lower for x in [
        "office ", "click-to-run", " mui ", " setup metadata",
        "office shared", "office 16", "office push notification",
        "excel", "word", "powerpoint", "outlook", "onedrivesetup",
    ]):
        return True

    # Skip Microsoft internal utilities
    if any(x in name_lower for x in [
        "game assist", "cross device", "copilot", "portal",
        "power automate", "quick assist", "print3d",
        "glanceby", "intel graphics experience", "personal gateway",
    ]):
        return True

    # Skip evaluation/trial versions
    if "evaluation" in name_lower or "trial" in name_lower:
        return True

    return False


def get_store_apps():
    software_list = []

    SKIP_PREFIXES = {
        "microsoft.windows.", "microsoft.ui.", "microsoft.net.",
        "microsoft.vclibrary", "microsoft.services.", "microsoft.directx",
        "microsoft.designtool", "microsoft.advertising", "microsoft.xbox",
        "microsoft.getstarted", "microsoft.gethelp", "microsoft.people",
        "microsoft.wallet", "microsoft.storepurchase", "microsoft.desktopappinstaller",
        "microsoft.webp", "microsoft.heif", "microsoft.hevc", "microsoft.raw",
        "microsoft.webmedia", "microsoft.vp9video", "microsoft.av1video",
        "microsoft.mpeg2video", "microsoft.screenSketch", "microsoft.paint",
        "microsoft.windowscamera", "microsoft.windowsmaps", "microsoft.windowsalarms",
        "microsoft.windowscalculator", "microsoft.windowscommunicationsapps",
        "microsoft.windowsfeedbackhub", "microsoft.windowsnotepad",
        "microsoft.windowssoundrecorder", "microsoft.windowsstore",
        "microsoft.windowsterminal", "microsoft.zunemusic", "microsoft.zunevideo",
        "microsoft.bingweather", "microsoft.bingnews", "microsoft.bingfinance",
        "microsoft.microsoftsolitairecollection", "microsoft.microsoftstickynotes",
        "microsoft.photos", "microsoft.todos", "microsoft.family",
        "microsoft.cortana", "microsoft.onedrive", "microsoft.yourphone",
        "microsoft.phonelinkstub", "microsoft.549981c3f5f10", "microsoft.mspaint",
        "microsoft.powershell", "microsoft.screensketch", "inputapp",
        "1527c705-839a-4832-9118", "c5e2524a-ea46-4f67-841f",
        "e2a4f912-2574-4a75-9bb0", "f46d4000-fd22-4db4-ac8e",
        "microsoft.asynctextservice", "microsoft.ecapp", "apprex.",
        "realtekcontrolcenter", "nvidia.", "lenovoinc.", "lenovokbapp",
        "dolbylaboratories.",
    }

    KNOWN_APPS = {
        "spotifyab.spotifymusic": "Spotify", "5319275a.whatsapp": "WhatsApp",
        "facebook.instagram": "Instagram", "facebook.facebook": "Facebook",
        "4df9e0f8.netflix": "Netflix", "9426micro.twitter": "Twitter",
        "tiktok.tiktok": "TikTok", "telegramfz": "Telegram",
        "discord": "Discord", "slack": "Slack", "zoom.zoom": "Zoom",
        "clipchamp": "Clipchamp", "canva": "Canva", "amazon": "Amazon",
        "disney": "Disney+", "primevideo": "Prime Video", "linkedin": "LinkedIn",
        "todoist": "Todoist", "notion": "Notion", "viber": "Viber",
        "line": "LINE", "itunestoday": "iTunes",
    }

    ps_cmd = (
        'Get-AppxPackage | Where-Object {$_.IsFramework -eq $false -and $_.SignatureKind -ne \"System\"} '
        '| Select-Object Name, @{N=\"Version\";E={$_.Version}}, Publisher '
        '| ConvertTo-Json -Compress'
    )

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            return []

        data = json.loads(result.stdout)
        if isinstance(data, dict):
            data = [data]

        for app in data:
            raw_name = app.get("Name", "") or ""
            version = app.get("Version", "") or ""
            publisher_raw = app.get("Publisher", "") or ""

            if not raw_name:
                continue

            name_lower = raw_name.lower()

            if any(name_lower.startswith(prefix) for prefix in SKIP_PREFIXES):
                continue

            friendly = None
            for fragment, label in KNOWN_APPS.items():
                if fragment in name_lower:
                    friendly = label
                    break

            if not friendly:
                parts = raw_name.split(".")
                tail = parts[-1] if len(parts) > 1 else parts[0]
                friendly = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', tail)
                if not friendly.strip():
                    continue

            publisher = publisher_raw
            if publisher.startswith("CN="):
                publisher = publisher.split(",")[0].replace("CN=", "")

            software_list.append({
                "Software": friendly.strip(),
                "Version": str(version).strip(),
                "Publisher": publisher.strip(),
                "InstallDate": "",
            })

    except (subprocess.TimeoutExpired, json.JSONDecodeError, FileNotFoundError):
        pass

    return software_list


def deduplicate_and_sort(software_list):
    seen = set()
    unique = []
    for entry in software_list:
        key = entry["Software"].lower()
        if key not in seen:
            seen.add(key)
            unique.append(entry)

    unique.sort(key=lambda x: x["Software"].lower())
    return unique


def extract_pdf_table(pdf_path):
    names = []
    sw_idx = None  # remember column index across pages
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                header = table[0]
                start = 1  # default: skip header row

                if sw_idx is None:
                    # First page: find the "Software" column
                    for i, col in enumerate(header):
                        col_name = (col or "").strip().lower()
                        if "software" in col_name:
                            sw_idx = i
                            break
                    if sw_idx is None:
                        sw_idx = 1 if len(header) > 1 else 0
                else:
                    # Continuation page: row 0 is data, not a header
                    start = 0

                for row in table[start:]:
                    if sw_idx < len(row):
                        val = (row[sw_idx] or "").strip().lower()
                        if val:
                            names.append(val)
    return names


PAID_SOFTWARE = [
    "mobaxterm",
    "termius",
    "docker desktop",
    "irfanview",
    "anaconda",
    "teracopy",
    "sublime",
    "windsurf",
    "warp",
    "burp suite professional",
    "laragon",
    "xshell",
    "sdk arm additions",
    "sdk arm redistributables",
    "vs script debugging common",
]

APPROVED_LIST = [
    "putty",
    "postman",
    "vscode",
    "mysql workbench",
    "python",
    "mremoteng",
    "notepad ++",
    "pycharm",
    "node.js",
    "mysql server",
    "mysql shell",
    "android studio",
    "dbeaver",
    "heidi sql",
    "mongo db compass",
    "mongo db tools",
    "win scp",
    "vs code",
    "cursor",
    "tabby",
    "wps office",
    "go",
    "git",
    "postgres",
    "oracle vm virtual box",
    "ruby",
    "xampp",
    "nvm",
    "vim",
    "pdf sam enhanced",
    "mqtt explorer",
    "redis insight",
    "eclipse",
    "cygwin64",
    "ireport",
    "mariadb",
    "netbeans",
    "soapui",
    "tortoisegit",
    "winmerge",
    "microsoft visual studio code",
    "wsl kali linux",
    "kitty",
    "pycharm community edition",
    "pgadmin",
    "obs",
    "slack",
    "filezilla",
    "telegram",
    "minikube",
    "neo4j desktop",
    "zed attack proxy",
    "nmap",
    "ncap",
    "obsidian",
    "workbench",
    "sql developer",
    "zenmap",
    "libre",
    "sqlmap",
    "searchsploit",
    "feroxbuster",
    "bitwarden",
    "intellij idea community",
    "greenshot",
    "owasp",
    "teltonika com driver",
    "telematic configuration tool",
    "draw.io",
    "microsoft visual c++",
    "telegram desktop",
    "obs studio",
    "python 3",
    "pip",
    "flir ip config",
    "tesseract ocr",
    "oracle jinitiator",
    "snc client encryption",
    "7-zip",
    "adobe refresh manager",
    "apache directory studio",
    "application verifier",
    "arduino",
    "atom",
    "blender",
    "brave",
    "clickonce bootstrapper",
    "cmder",
    "discord",
    "eclipse mosquitto",
    "entity framework",
    "gns3",
    "gnuwin32",
    "go programming language",
    "google earth pro",
    "gpl ghostscript",
    "handbrake",
    "heidi",
    "helm",
    "hxd hex editor",
    "hydra",
    "icecap_collection",
    "inkscape",
    "intellij idea",
    "ivi shared components",
    "keepass",
    "keepassxc",
    "kicad",
    "kits configuration installer",
    "laragon",
    "mediatek sp driver",
    "microsoft .net framework",
    "microsoft .net host",
    "microsoft .net targeting pack",
    "microsoft gameinput",
    "microsoft netstandard sdk",
    "microsoft testplatform sdk",
    "microsoft windows desktop targeting pack",
    "miniforge",
    "mobile broadband hl service",
    "msi development tools",
    "mtk usb all",
    "mysql installer",
    "npgsql",
    "numbers",
    "ollama",
    "onlyoffice",
    "open vpn",
    "openlens",
    "opera",
    "oracle sql developer",
    "oracle virtualbox",
    "owasp zap",
    "pdf24 creator",
    "pdfsam basic",
    "photopea",
    "qgis",
    "quranflash",
    "r for windows",
    "remmina",
    "rstudio",
    "sd card formatter",
    "sdk arm64",
    "sdt commonssense tool",
    "soap ui",
    "sound lock",
    "sound recorder",
    "sqlcmd tools",
    "starship",
    "sticky note",
    "universal crt extension sdk",
    "universal crt headers",
    "universal crt redistributable",
    "universal general midi",
    "vagrant",
    "vcpp_crt",
    "visa.net shared components",
    "visual c++ library",
    "visual studio build tools 2022",
    "vlc media player",
    "vs immersive activate helper",
    "vs_blendmsi",
    "vs_clickoncebootstrappermsires",
    "vs_communitymsires",
    "vs_coreeditorfonts",
    "vs_devenvsharedmsi",
    "vs_filehandler_amd64",
    "vs_filetracker_singleton",
    "vs_githubprotocolhandlermsi",
    "vs_graphics_singletonx86",
    "vs_minshellinteropsharedmsi",
    "vs_sqlclickoncebootstrappermsi",
    "vs_tipsmsi",
    "vs_vswebprotocolselectormsi",
    "vscodium",
    "waterfox",
    "webtablet fb plugin",
    "wimgapi",
    "winappdeploy",
    "winbox64",
    "windows app certification kit",
    "windows iot extension sdk",
    "windows mobile extension sdk",
    "windows sdk",
    "windows_toolscorepkg",
    "winrt intellisense iot",
    "winrt intellisense ppi",
    "wireshark targeting pack",
    "wpt redistributables",
    "wptx64",
    "zoom",
    "sharex",
    "wacom tablet driver",
    "db browser for sqlite",
    "microsoft powerbi desktop",
    "power automate for desktop",
    "whatsapp",
    "winscp",
    "java 8",
    "openjdk",
    "wsl ubuntu",
    "github desktop",
    "ffmpeg",
    "audacity",
    "miktex",
    "texstudio",
    "laslook",
    "mobaxterm",
    "termius",
    "docker desktop",
    "irfanview",
    "anaconda",
    "teracopy",
    "sublime",
    "windsurf",
    "warp",
    "burp suite professional",
    "xshell",
    "sdk arm additions",
    "sdk arm redistributables",
    "vs script debugging common",
]

def build_official_list():
    entries = []
    seen = set()

    # Approved: use hardcoded list
    for name in APPROVED_LIST:
        clean = name.strip().lower()
        if clean and clean not in seen:
            entries.append((clean, "Allowed"))
            seen.add(clean)

    # Not Allowed: still read from PDF
    for f in NOT_APPROVED_PDFS:
        if os.path.isfile(f):
            for name in extract_pdf_table(f):
                if name not in seen:
                    entries.append((name, "Not Allowed"))
                    seen.add(name)

    return entries


def clean_name(name):
    if name is None:
        return ""
    name = str(name).lower()
    name = re.sub(r"\d+\.\d+[\.\d]*", "", name)
    name = re.sub(r"\(x64\)|\(64-bit\)|\(32-bit\)", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[^\w\s]", " ", name)
    return " ".join(name.split()).strip()


def run_check_local(csv_bytes):
    csv_rows = []
    for line in csv_bytes.decode("utf-8").splitlines(keepends=True):
        if not line.startswith("#"):
            csv_rows.append(line)

    user_reader = csv.DictReader(io.StringIO("".join(csv_rows)))

    official = build_official_list()
    official_names = [name for name, _ in official]
    official_status = [status for _, status in official]
    official_clean = [clean_name(name) for name in official_names]

    results = []
    counts = {"Allowed": 0, "Not Allowed": 0, "Not Found": 0}

    ignore_count = 0
    for row in user_reader:
        sw_name = str(row.get("Software") or "")

        # Skip if in ignore list (substring match) or too short to be real software
        if any(ignored.lower() in sw_name.lower() for ignored in IGNORE_LIST):
            ignore_count += 1
            continue
        if len(sw_name.strip()) <= 2:
            ignore_count += 1
            continue

        sw_clean = clean_name(sw_name)
        matched = ""
        status = "Not Found"

        for idx, off_clean in enumerate(official_clean):
            if not off_clean or not sw_clean:
                continue
            if len(sw_clean) < 3 or len(off_clean) < 3:
                continue
            if sw_clean in off_clean or off_clean in sw_clean:
                matched = official_names[idx]
                status = official_status[idx]
                break

        # Check if matched software is paid
        remark = ""
        if status == "Allowed" and matched:
            if any(p in matched for p in PAID_SOFTWARE):
                remark = "Paid software. Please contact IT unit."

        results.append({
            "software": sw_name,
            "status": status,
            "matched": matched,
            "remark": remark,
        })
        counts[status] = counts.get(status, 0) + 1

    results.sort(key=lambda r: (["Not Allowed", "Not Found", "Allowed"].index(r["status"]), r["software"].lower()))
    return results, counts


def main():
    # Force UTF-8 output to handle special characters in all terminals
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    try:
        hostname = socket.gethostname()
        username = getpass.getuser()
        scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        print("\n" + "="*80)
        print(f"  SOFTWARE COMPLIANCE SCAN V1.1")
        print("="*80)
        print(f"  Hostname: {hostname}")
        print(f"  Username: {username}")
        print(f"  Time: {scan_time}")
        print("="*80 + "\n")

        # Step 1: SCAN
        print("  [STEP 1] Scanning installed software...\n")
        registry_apps = get_installed_software()
        store_apps = get_store_apps()
        software = deduplicate_and_sort(registry_apps + store_apps)
        print(f"  [OK] Found installed programs.\n")

        # Step 2: Create CSV in memory for compliance check
        csv_output = io.StringIO()
        csv_output.write(f"# Hostname: {hostname}\n")
        csv_output.write(f"# Username: {username}\n")
        csv_output.write(f"# Scan Time: {scan_time}\n")
        writer = csv.DictWriter(csv_output, fieldnames=["Software", "Version", "Publisher", "InstallDate"])
        writer.writeheader()
        writer.writerows(software)
        csv_bytes = csv_output.getvalue().encode("utf-8")

        # Step 3: COMPARE with official list
        print("  [STEP 2] Comparing against compliance lists...\n")
        results, counts = run_check_local(csv_bytes)

        # Step 4: SAVE compliance report to Downloads
        safe_hostname = hostname.replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"compliance_report_{safe_hostname}_{timestamp}.csv"
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(downloads_dir, exist_ok=True)
        output_path = os.path.join(downloads_dir, filename)

        with open(output_path, "w", newline="", encoding="utf-8") as f:
            f.write(f"Software Compliance Report\n")
            f.write(f"Hostname: {hostname}\n")
            f.write(f"Username: {username}\n")
            f.write(f"Scan Time: {scan_time}\n")
            f.write(f"Total Checked: {sum(counts.values())}\n")
            f.write(f"\n")

            w = csv.writer(f)

            allowed = [r for r in results if r["status"] == "Allowed"]
            not_allowed = [r for r in results if r["status"] == "Not Allowed"]
            unknown = [r for r in results if r["status"] == "Not Found"]

            f.write(f"NOT ALLOWED ({len(not_allowed)}) - Please uninstall immediately or contact IT unit\n")
            w.writerow(["Software", "Matched With"])
            for r in not_allowed:
                w.writerow([r["software"], r["matched"]])
            f.write(f"\n")

            f.write(f"ALLOWED ({len(allowed)})\n")
            w.writerow(["Software", "Matched With", "Remark"])
            for r in allowed:
                w.writerow([r["software"], r["matched"], r.get("remark", "")])
            f.write(f"\n")

            f.write(f"UNKNOWN ({len(unknown)})\n")
            w.writerow(["Software"])
            for r in unknown:
                w.writerow([r["software"]])

        print(f"  [OK] Saved to: {output_path}\n")
        
        # Step 4: DISPLAY RESULTS
        display_results(results, counts, hostname)
        
        print("\n" + "="*80)
        print("  SCAN COMPLETE")
        print("="*80)
        print(f"  CSV saved to: {output_path}")
        print("="*80 + "\n")
        
        return output_path

    except Exception as e:
        print(f"\n  [ERROR] {str(e)}", flush=True)
        import traceback
        traceback.print_exc()
        return None


def display_results(results, counts, hostname):
    """Display compliance results in a formatted table."""
    
    # Group results by status
    allowed = [r for r in results if r["status"] == "Allowed"]
    not_allowed = [r for r in results if r["status"] == "Not Allowed"]
    not_found = [r for r in results if r["status"] == "Not Found"]

    # Print summary
    print(f"  ")
    print(f"  +-- COMPLIANCE SUMMARY -----------------------+")
    print(f"  | [OK]  Allowed:      {counts['Allowed']:>4} programs            |")
    print(f"  | [!!]  Not Allowed:  {counts['Not Allowed']:>4} programs            |")
    print(f"  | [??]  Not Found:    {counts.get('Not Found', 0):>4} programs            |")
    print(f"  +---------------------------------------------+\n")

    # Print detailed results if there are issues
    if not_allowed:
        print(f"  === NOT ALLOWED ({len(not_allowed)}) ========================================")
        for r in not_allowed:
            print(f"  | {r['software'][:60]}")
            print(f"  |   Matched: {r['matched'][:55]}")
        print(f"  |")
        print(f"  | NOTE: Please uninstall immediately or contact IT unit.")
        print(f"  ================================================================\n")

    if allowed:
        print(f"  === ALLOWED ({len(allowed)}) ===========================================")
        for r in allowed:
            name = r['software'][:60]
            if r.get('remark'):
                name += " (paid software)"
            print(f"  | {name}")
            print(f"  |   Matched: {r['matched'][:55]}")
        print(f"  ================================================================\n")

    if not_found:
        print(f"  === UNKNOWN ({len(not_found)}) ============================================")
        for r in not_found[:20]:  # Show first 20
            print(f"  | {r['software'][:60]}")
        if len(not_found) > 20:
            print(f"  | ... and {len(not_found) - 20} more ...")
        print(f"  ================================================================\n")

    if allowed and not not_allowed and not not_found:
        print(f"  [OK] All programs are compliant!\n")


if __name__ == "__main__":
    main()
    input("\n  Press Enter to close...")

