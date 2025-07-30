import os
import shutil
import sys
import tempfile
import datetime
import platform
import subprocess
import py7zr
import concurrent.futures

EXCLUDE_EXTENSIONS = ('.dll', '.exe', '.bin', '.msi', '.dat', '.rar', '.gz', '.cab')
MAX_FILE_SIZE = 8 * 1024 * 1024  # 8 MB


def get_windows_temp_path():
    for path in [r"C:\Windows\Temp", r"C:\Temp"]:
        try:
            os.makedirs(path, exist_ok=True)
            return path
        except:
            continue
    return tempfile.gettempdir()


def export_event_logs(output_dir):
    logs_to_export = ["Application", "System", "Security"]
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    logs_dir = os.path.join(output_dir, "EventLogs")
    os.makedirs(logs_dir, exist_ok=True)

    for log in logs_to_export:
        safe_name = log.replace("/", "_")
        path = os.path.join(logs_dir, f"{safe_name}_{timestamp}.evtx")
        try:
            subprocess.run(["wevtutil", "epl", log, path],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        except:
            continue
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
        os.path.expandvars(r"%ProgramData%\GetSupportService_LOGICnow"),
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


def should_ignore(file_path):
    if not os.path.isfile(file_path):
        return False
    ext = os.path.splitext(file_path)[1].lower()
    return ext in EXCLUDE_EXTENSIONS or os.path.getsize(file_path) > MAX_FILE_SIZE


def process_category(category, paths, root_output):
    category_folder = os.path.join(root_output, category)
    os.makedirs(category_folder, exist_ok=True)
    copied = False

    for path in paths:
        try:
            if callable(path):
                result_path = path(tempfile.mkdtemp())
                if os.path.exists(result_path) and os.listdir(result_path):
                    dest = os.path.join(category_folder, "EventLogs")
                    shutil.copytree(result_path, dest, dirs_exist_ok=True)
                    copied = True
                continue

            if os.path.isfile(path):
                if not should_ignore(path):
                    drive, relative = os.path.splitdrive(path)
                    relative = relative.lstrip("\\/")
                    dest_path = os.path.join(category_folder, relative)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    shutil.copy2(path, dest_path)
                    copied = True

            elif os.path.isdir(path):
                for root, dirs, files in os.walk(path):
                    for file in files:
                        full_path = os.path.join(root, file)
                        if should_ignore(full_path):
                            continue
                        drive, relative = os.path.splitdrive(full_path)
                        relative = relative.lstrip("\\/")
                        dest_path = os.path.join(category_folder, relative)
                        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                        shutil.copy2(full_path, dest_path)
                        copied = True
        except:
            continue
    return copied


def copy_all_categories(destination_folder):
    copied_any = False
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(process_category, category, paths, destination_folder)
            for category, paths in categories.items()
        ]
        for future in concurrent.futures.as_completed(futures):
            if future.result():
                copied_any = True
    return copied_any


def create_7z_archive(source_folder, archive_path):
    try:
        with py7zr.SevenZipFile(archive_path, 'w') as archive:
            archive.writeall(source_folder, arcname='.')
        return True
    except:
        return False


def run_silent():
    temp_dir = tempfile.mkdtemp()
    output_dir = get_windows_temp_path()
    os.makedirs(output_dir, exist_ok=True)

    archive_name = generate_archive_name()
    archive_path = os.path.join(output_dir, archive_name)

    if copy_all_categories(temp_dir):
        if create_7z_archive(temp_dir, archive_path):
            shutil.rmtree(temp_dir, ignore_errors=True)
            sys.exit(0)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)
            sys.exit(1)
    else:
        shutil.rmtree(temp_dir, ignore_errors=True)
        sys.exit(2)


if __name__ == "__main__":
    run_silent()
