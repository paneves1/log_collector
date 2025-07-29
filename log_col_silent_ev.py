import os
import shutil
import sys
import tempfile
import datetime
import platform
import pythoncom
import win32com.client
import py7zr
import subprocess

EXCLUDE_EXTENSIONS = ('.dll', '.exe', '.bin', '.msi', '.dat', '.rar', '.gz','cab')

# Function to get the Windows temporary path
# This function tries to use the primary temp path and falls back to a secondary one if it fails

def get_windows_temp_path():
    primary = r"C:\Windows\Temp"
    fallback = r"C:\Temp"

    try:
        os.makedirs(primary, exist_ok=True)
        return primary
    except Exception:
        os.makedirs(fallback, exist_ok=True)
        return fallback



def export_event_logs(output_dir):
    logs_to_export = [
        "Application",
        "System",
        "Security"
    ]
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    logs_dir = os.path.join(output_dir, "EventLogs")
    os.makedirs(logs_dir, exist_ok=True)

    for log_name in logs_to_export:
        sanitized_name = log_name.replace("/", "_")
        file_path = os.path.join(logs_dir, f"{sanitized_name}_{timestamp}.evtx")
        try:
            subprocess.run(
                ["wevtutil", "epl", log_name, file_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True
            )
        except subprocess.CalledProcessError:
            pass

    return logs_dir

categories = {
    "Automation Manager": [
        os.path.expandvars(r"C:\Program Files (x86)\N-able Technologies\AutomationManager\logs\\"),
        os.path.expandvars(r"C:\Program Files (x86)\Advanced Monitoring Agent\scriptrunner\\"),
        os.path.expandvars(r"C:\ProgramData\N-able Technologies\AutomationManager\log\\"),
        os.path.expandvars(r"C:\ProgramData\N-able Technologies\AutomationManager\scripts\\"),
    ],
    "MSP Core": [
        os.path.expandvars(r"C:\Program Files (x86)\Msp Agent\\"),
    ],
    "Vulnerability Management": [
        os.path.expandvars(r"C:\Program Files (x86)\Msp Agent\Components\software-scanner\\"),
        os.path.expandvars(r"C:\ProgramData\N-able Technologies\Vulnerability Management\logs\\"),
    ],
    "Take Control Console": [
        os.path.expandvars(r"%LOCALAPPDATA%\BeAnywhere Support Express\Console\Logs\\"),
    ],
    "Take Control StandAlone Agent": [
        os.path.expandvars(r"%ALLUSERSPROFILE%\GetSupportService\Logs\\"),
    ],
    "N-sight Agent": [
        os.path.expandvars(r"C:\Program Files (x86)\Advanced Monitoring Agent"),
        os.path.expandvars(r"C:\Program Files (x86)\Advanced Monitoring Agent GP"),
        os.path.expandvars(r"%ProgramData%\MspPlatform\PME\log"),
        os.path.expandvars(r"%ProgramData%\MspPlatform\FileCacheServiceAgent\log"),
        os.path.expandvars(r"%ProgramData%\MspPlatform\PME.Agent.PmeService\log"),
        os.path.expandvars(r"%ProgramData%\MspPlatform\RequestHandlerAgent\log"),
        os.path.expandvars(r"%ProgramData%\AdvancedMonitoringAgentWebProtection"),
        os.path.expandvars(r"%ProgramData%\AdvancedMonitoringAgentNetworkManagement"),
    ],
    "Take Control Viewer": [
        os.path.expandvars(r"%LOCALAPPDATA%\Take Control Viewer\Logs\\"),
    ],
    "Event Viewer Logs": [
        export_event_logs
    ],
}

def generate_archive_name():
    hostname = platform.node()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"N-Able_Logs_{hostname}_{timestamp}.7z"

def copy_selected_items(destination_folder):
    files_copied = False

    for category, paths in categories.items():
        temp_subfolder = tempfile.mkdtemp()
        copied_any = False

        for path in paths:
            try:
                if callable(path):  # Export Event logs
                    result_path = path(temp_subfolder)
                    if os.path.exists(result_path) and os.listdir(result_path):
                        copied_any = True
                        files_copied = True
                elif os.path.isfile(path):
                    if not path.lower().endswith(EXCLUDE_EXTENSIONS):
                        os.makedirs(temp_subfolder, exist_ok=True)
                        shutil.copy2(path, temp_subfolder)
                        copied_any = True
                        files_copied = True
                elif os.path.isdir(path) and os.listdir(path):
                    dest = os.path.join(temp_subfolder, os.path.basename(path))
                    os.makedirs(os.path.dirname(dest), exist_ok=True)

                    def ignore_files(_, filenames):
                        return [f for f in filenames if f.lower().endswith(EXCLUDE_EXTENSIONS)]

                    shutil.copytree(path, dest, dirs_exist_ok=True, ignore=ignore_files)
                    copied_any = True
                    files_copied = True
            except:
                pass

        if copied_any:
            final_category_folder = os.path.join(destination_folder, category)
            os.makedirs(final_category_folder, exist_ok=True)
            for item in os.listdir(temp_subfolder):
                shutil.move(os.path.join(temp_subfolder, item), final_category_folder)

        shutil.rmtree(temp_subfolder, ignore_errors=True)

    return files_copied

def create_7z_archive(source_folder, archive_path):
    try:
        with py7zr.SevenZipFile(archive_path, 'w') as archive:
            archive.writeall(source_folder, arcname='.')
        return True
    except:
        return False

def run_silent():
    temp_dir = tempfile.mkdtemp()
    desktop_dir = get_windows_temp_path()  # Temporary directory for the archive
    os.makedirs(desktop_dir, exist_ok=True)

    archive_name = generate_archive_name()
    archive_path = os.path.join(desktop_dir, archive_name)

    files_copied = copy_selected_items(temp_dir)

    if files_copied:
        if create_7z_archive(temp_dir, archive_path):
            shutil.rmtree(temp_dir, ignore_errors=True)
            sys.exit(0)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)
            sys.exit(1)
    else:
        shutil.rmtree(temp_dir, ignore_errors=True)
        sys.exit(2)

        sys.exit(2)

if __name__ == "__main__":
    run_silent()
