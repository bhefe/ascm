import os
import sys

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

from scan_software import get_installed_software, get_store_apps, deduplicate_and_sort
import socket
import getpass
import csv
from datetime import datetime


def scan():
    print("=" * 50)
    print("  SOFTWARE COMPLIANCE CHECK")
    print("=" * 50)

    hostname = socket.gethostname()
    username = getpass.getuser()
    scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"\n  Scanning {hostname}...")
    registry_apps = get_installed_software()
    store_apps = get_store_apps()
    software = deduplicate_and_sort(registry_apps + store_apps)
    print(f"  Found {len(software)} installed programs.")

    safe_hostname = hostname.replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"installed_software_{safe_hostname}_{timestamp}.csv"
    output_path = os.path.join(SCRIPT_DIR, filename)

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        f.write(f"# Hostname: {hostname}\n")
        f.write(f"# Username: {username}\n")
        f.write(f"# Scan Time: {scan_time}\n")
        writer = csv.DictWriter(f, fieldnames=["Software", "Version", "Publisher", "InstallDate"])
        writer.writeheader()
        writer.writerows(software)

    return output_path


from check_upload import process_upload


def check(csv_path):
    print("\n  Running compliance check...")
    return process_upload(csv_path)


if __name__ == "__main__":
    csv_path = scan()
    report_path = check(csv_path)

    print()
    print("=" * 50)
    print(f"  Report: {os.path.basename(report_path)}")
    print("=" * 50)

    input("\n  Press Enter to close...")
