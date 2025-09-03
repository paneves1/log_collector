import os
import shutil
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
import datetime
import sys
import ctypes
import platform
import tempfile
import pythoncom
import win32com.client
import py7zr
from functools import partial

EXCLUDE_EXTENSIONS = tuple(ext.lower() for ext in ('.dll', '.exe', '.bin', '.msi', '.dat', '.rar', '.gz','cab'))
MAX_FILE_SIZE = 8 * 1024 * 1024  # 8 MB in bytes

# --- Get user's Desktop path ---
def get_desktop_path():
    pythoncom.CoInitialize()
    shell = win32com.client.Dispatch("WScript.Shell")
    return shell.SpecialFolders("Desktop")

# --- Setup Logging ---
desktop_path = get_desktop_path()
log_file = os.path.join(desktop_path, "N-Able_LogCollector.log")
if not logging.getLogger().hasHandlers():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='a', encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

# --- Global control flags ---
is_collecting = False
current_temp_dir = None

# --- Check for administrator privileges ---
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# --- Exclusion helper ---
def is_excluded_file(file_path):
    return file_path.lower().endswith(EXCLUDE_EXTENSIONS)

# --- Log categories and paths ---
categories = {
    "Automation Manager": [
        os.path.expandvars(r"C:\Program Files (x86)\N-able Technologies\AutomationManager\logs"),
        os.path.expandvars(r"C:\Program Files (x86)\Advanced Monitoring Agent\scriptrunner"),
        os.path.expandvars(r"C:\ProgramData\N-able Technologies\AutomationManager\log"),
        os.path.expandvars(r"C:\ProgramData\N-able Technologies\AutomationManager\scripts"),
    ],
    "MSP Core": [
        os.path.expandvars(r"C:\Program Files (x86)\Msp Agent\logs"),
    ],
    "Vulnerability Management": [
        os.path.expandvars(r"C:\Program Files (x86)\Msp Agent\Components\software-scanner"),
        os.path.expandvars(r"C:\ProgramData\N-able Technologies\Vulnerability Management\logs"),
    ],
    "Take Control Console": [
        os.path.expandvars(r"%LOCALAPPDATA%\BeAnywhere Support Express\Console\Logs"),
    ],
    "Take Control StandAlone Agent": [
        os.path.expandvars(r"%ALLUSERSPROFILE%\GetSupportService\Logs"),
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
        os.path.expandvars(r"%ProgramData%\GetSupportService_LOGICnow"),
    ],
    "Event Logs": [],
    "Take Control Viewer": [
        os.path.expandvars(r"%LOCALAPPDATA%\Take Control Viewer\Logs"),
    ],
}

# --- Generate archive name ---
def generate_archive_name():
    hostname = platform.node()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"N-Able_Logs_{hostname}_{timestamp}.7z"

# --- Copy files/folders for a category ---
def copy_selected_items_for_category(category, destination_folder, exclude_large_files=False):
    paths = categories.get(category, [])
    copied = False
    temp_subfolder = tempfile.mkdtemp()

    for path in paths:
        try:
            if os.path.isfile(path):
                if is_excluded_file(path):
                    continue
                if exclude_large_files and os.path.getsize(path) > MAX_FILE_SIZE:
                    logging.info(f"[{category}] Skipped large file: {path}")
                    continue
                os.makedirs(temp_subfolder, exist_ok=True)
                shutil.copy2(path, temp_subfolder)
                logging.info(f"[{category}] Copied file: {path}")
                copied = True
            elif os.path.isdir(path):
                def ignore_func(dir_path, dir_names):
                    ignored = []

                    # Exclude files by extension
                    for f in list(dir_names):
                        if is_excluded_file(f):
                            ignored.append(f)
                    return ignored
                
                # Preserve full path relative to the root of the drive (e.g., from C:\)
                drive, _ = os.path.splitdrive(path)
                rel_path = os.path.relpath(path, start=drive + os.sep)
                dest = os.path.join(temp_subfolder, drive.strip(":") + os.sep + rel_path)
                shutil.copytree(path, dest, dirs_exist_ok=True, ignore=ignore_func)
                
                if exclude_large_files:
                    for root, _, files in os.walk(dest):
                        for f in files:
                            fpath = os.path.join(root, f)
                            try:
                                if os.path.getsize(fpath) > MAX_FILE_SIZE:
                                    os.remove(fpath)
                                    logging.info(f"[{category}] Removed large file: {fpath}")
                            except Exception as e:
                                logging.warning(f"Error checking/removing file {fpath}: {e}")
                
                logging.info(f"[{category}] Copied folder: {path} to {dest}")
                copied = True
        except Exception as e:
            logging.warning(f"[{category}] Could not copy {path}: {e}")

    if copied:
        safe_category = category.replace(" ", "_").replace("-", "_")
        final_folder = os.path.join(destination_folder, safe_category)
        os.makedirs(final_folder, exist_ok=True)
        for item in os.listdir(temp_subfolder):
            shutil.move(os.path.join(temp_subfolder, item), final_folder)

    shutil.rmtree(temp_subfolder, ignore_errors=True)
    return copied

# --- Export Windows Event Logs ---
def export_event_logs(destination_folder):
    try:
        temp_dir = tempfile.mkdtemp()
        logs = ["System", "Application", "Security"]
        files_exported = 0

        for logname in logs:
            outfile = os.path.join(temp_dir, f"{logname}.evtx")
            result = os.system(f'wevtutil epl {logname} "{outfile}"')
            if result != 0:
                logging.error(f"Failed to export log: {logname}")
                continue
            if os.path.isfile(outfile):
                eventlogs_dir = os.path.join(destination_folder, "EventLogs")
                os.makedirs(eventlogs_dir, exist_ok=True)
                shutil.move(outfile, os.path.join(eventlogs_dir, f"{logname}.evtx"))
                logging.info(f"Exported event log: {logname}")
                files_exported += 1

        shutil.rmtree(temp_dir)
        return files_exported > 0
    except Exception as e:
        logging.error(f"Failed to export Event Logs: {e}", exc_info=True)
        messagebox.showerror("Error", f"Fail to export Event Logs: {e}")
        return False

# --- Create 7z archive ---
def create_7z_archive(source_folder, archive_path):
    try:
        if not has_valid_files(source_folder):
            logging.warning("Nothing to archive: source folder is empty or no valid files.")
            return False
        with py7zr.SevenZipFile(archive_path, 'w', filters=[{'id': py7zr.FILTER_LZMA2, 'preset': 1}]) as archive:
            archive.writeall(source_folder, arcname=platform.node())
        return True
    except Exception as e:
        logging.error(f"Failed to create archive: {e}")
        return False

# --- Check if folder has valid files ---
def has_valid_files(folder):
    for root, _, files in os.walk(folder):
        if any(not f.lower().endswith(EXCLUDE_EXTENSIONS) for f in files):
            return True
    return False

# --- Preload N-sight Agent ---
preloaded_n_sight_dir = tempfile.mkdtemp()

def preload_n_sight_agent():
    logging.info("Preloading N-sight Agent logs...")
    copy_selected_items_for_category("N-sight Agent", preloaded_n_sight_dir)
    logging.info("N-sight Agent preload complete.")

# --- Collect logs ---
def collect_logs(progress_bar, progress_label, checkboxes):
    global is_collecting, current_temp_dir
    is_collecting = True
    progress_bar["value"] = 0
    progress_label.config(text="0%")
    selected = [cat for cat, var in checkboxes.items() if var.get()]
    if not selected:
        messagebox.showinfo("Info", "No category selected.")
        is_collecting = False
        return

    total = len(selected)
    step = 100 / total
    current_temp_dir = tempfile.mkdtemp()
    temp_dir = current_temp_dir
    archive_path = os.path.join(desktop_path, generate_archive_name())

    logging.info(f"Selected categories: {selected}")
    logging.info(f"Temporary directory: {temp_dir}")
    files_copied = False

    for i, category in enumerate(selected):
        if category == "Event Logs":
            if export_event_logs(temp_dir):
                files_copied = True
        elif category == "N-sight Agent":
            preload_dest = os.path.join(temp_dir, "N_sight_Agent")
            os.makedirs(preload_dest, exist_ok=True)
            for item in os.listdir(preloaded_n_sight_dir):
                src = os.path.join(preloaded_n_sight_dir, item)
                dst = os.path.join(preload_dest, item)
                if os.path.isdir(src):
                    shutil.copytree(src, dst, dirs_exist_ok=True)
                else:
                    shutil.copy2(src, dst)
            if has_valid_files(preload_dest):
                files_copied = True
            else:
                logging.warning("N-sight Agent: No valid files found in preloaded data.")
        else:
            # Always exclude large files except for Event Logs (already handled)
            if copy_selected_items_for_category(category, temp_dir, exclude_large_files=True):
                files_copied = True

        percent = int((i + 1) * step)
        progress_bar.after(0, partial(progress_bar.config, value=percent))
        progress_label.after(0, partial(progress_label.config, text=f"{percent}%"))

    if not has_valid_files(temp_dir):
        messagebox.showinfo("Info", "No files or folders found. Archive not created.")
    else:
        if create_7z_archive(temp_dir, archive_path):
            messagebox.showinfo("Success", f"Archive created: {archive_path}")
        else:
            messagebox.showerror("Error", "Failed to create .7z archive")

    progress_bar.after(0, lambda: progress_bar.config(value=0))
    progress_label.after(0, lambda: progress_label.config(text="0%"))
    is_collecting = False
    current_temp_dir = None
    shutil.rmtree(temp_dir, ignore_errors=True)

# --- Start log collection in a thread ---
def start_collection(progress_bar, progress_label, checkboxes, collect_btn):
    collect_btn.config(state="disabled")
    def run():
        collect_logs(progress_bar, progress_label, checkboxes)
        progress_bar.after(0, lambda: collect_btn.config(state="normal"))
    threading.Thread(target=run, daemon=True).start()

# --- Create GUI ---
def create_gui():
    root = tk.Tk()
    root.title("N-Able Log Collector")
    root.geometry("320x370")

    def on_closing():
        global is_collecting, current_temp_dir
        if is_collecting:
            if messagebox.askyesno("Collecting logs", "Logs are still being collected. Do you want to force exit?"):
                logging.warning("User closed the application during log collection. Archive creation incomplete.")
                is_collecting = False
                if current_temp_dir and os.path.isdir(current_temp_dir):
                    shutil.rmtree(current_temp_dir, ignore_errors=True)
                root.destroy()
        else:
            if messagebox.askokcancel("Quit", "Do you want to exit the Log Collector?"):
                root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    admin_status = "Yes" if is_admin() else "No"
    tk.Label(root, text=f"Administrator privileges: {admin_status}", fg="green" if is_admin() else "red").grid(row=0, column=0, columnspan=2, pady=5)
    tk.Label(root, text="Select log categories to collect:").grid(row=1, column=0, columnspan=2, pady=5)

    checkboxes = {}
    for i, category in enumerate(categories):
        var = tk.BooleanVar()
        tk.Checkbutton(root, text=category, variable=var).grid(row=i+2, column=0, sticky="w", padx=20)
        checkboxes[category] = var

    progress_bar = ttk.Progressbar(root, length=300, mode="determinate")
    progress_bar.grid(row=len(categories)+2, column=0, columnspan=2, pady=10, padx=(10, 0))

    progress_label = tk.Label(root, text="0%")
    progress_label.grid(row=len(categories)+3, column=0, columnspan=2)

    collect_btn = tk.Button(root, text="Collect Logs")
    collect_btn.config(command=lambda: start_collection(progress_bar, progress_label, checkboxes, collect_btn))
    collect_btn.grid(row=len(categories)+4, column=0, columnspan=2, pady=10)

    threading.Thread(target=preload_n_sight_agent, daemon=True).start()

    root.mainloop()

# --- Main ---
if __name__ == "__main__":
    logging.info(f"Running with administrator privileges: {'Yes' if is_admin() else 'No'}")
    create_gui()
#

