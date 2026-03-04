import winreg
import csv
import os
import socket
import getpass
import subprocess
import json
import re
from datetime import datetime


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


def main():
    hostname = socket.gethostname()
    username = getpass.getuser()
    scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"  Scanning {hostname}...")

    registry_apps = get_installed_software()
    store_apps = get_store_apps()
    software = deduplicate_and_sort(registry_apps + store_apps)
    print(f"  Found {len(software)} installed programs.")

    safe_hostname = hostname.replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"installed_software_{safe_hostname}_{timestamp}.csv"

    # Always save to the user's Downloads folder
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads_dir, exist_ok=True)
    output_path = os.path.join(downloads_dir, filename)

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        f.write(f"# Hostname: {hostname}\n")
        f.write(f"# Username: {username}\n")
        f.write(f"# Scan Time: {scan_time}\n")
        writer = csv.DictWriter(f, fieldnames=["Software", "Version", "Publisher", "InstallDate"])
        writer.writeheader()
        writer.writerows(software)

    print(f"  Saved: {filename}")
    return output_path


if __name__ == "__main__":
    main()
    input("\n  Press Enter to close...")
