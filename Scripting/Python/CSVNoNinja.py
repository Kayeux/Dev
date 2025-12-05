import sys
import subprocess
import importlib.util
import ctypes
from ctypes import wintypes
import os
import getpass

#
# SECTION 1: Auto-install pandas if missing
#

def ensure_pandas_installed():
    """
    Check if 'pandas' can be imported. If not, attempts to install via PowerShell pip.
    Exits on error.
    """
    if importlib.util.find_spec("pandas") is None:
        print("Module 'pandas' not found. Attempting to install via pip...")
        try:
            subprocess.check_call([
                "powershell",
                "-Command",
                "pip install pandas"
            ])
        except subprocess.CalledProcessError as e:
            print(f"Failed to install 'pandas'. Error:\n{e}")
            sys.exit(1)

        # Verify install succeeded
        if importlib.util.find_spec("pandas") is None:
            print("pandas still not found after pip install. Exiting.")
            sys.exit(1)
        else:
            print("Successfully installed pandas.")
    else:
        print("Module 'pandas' is already installed.")


#
# SECTION 2: Define Windows constants & structures for file dialogs
#

MAX_PATH = 260

OFN_FILEMUSTEXIST      = 0x00001000
OFN_PATHMUSTEXIST      = 0x00000800
OFN_OVERWRITEPROMPT    = 0x00000002

# comdlg32 functions
GetOpenFileNameW = ctypes.windll.comdlg32.GetOpenFileNameW
GetSaveFileNameW = ctypes.windll.comdlg32.GetSaveFileNameW

# The OPENFILENAMEW structure as used by GetOpenFileNameW / GetSaveFileNameW
class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize",       wintypes.DWORD),
        ("hwndOwner",         wintypes.HWND),
        ("hInstance",         wintypes.HINSTANCE),
        ("lpstrFilter",       wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter",    wintypes.DWORD),
        ("nFilterIndex",      wintypes.DWORD),
        ("lpstrFile",         wintypes.LPWSTR),   # pointer to buffer
        ("nMaxFile",          wintypes.DWORD),
        ("lpstrFileTitle",    wintypes.LPWSTR),
        ("nMaxFileTitle",     wintypes.DWORD),
        ("lpstrInitialDir",   wintypes.LPCWSTR),
        ("lpstrTitle",        wintypes.LPCWSTR),
        ("Flags",             wintypes.DWORD),
        ("nFileOffset",       wintypes.WORD),
        ("nFileExtension",    wintypes.WORD),
        ("lpstrDefExt",       wintypes.LPCWSTR),
        ("lCustData",         ctypes.c_void_p),
        ("lpfnHook",          ctypes.c_void_p),
        ("lpTemplateName",    wintypes.LPCWSTR),
        ("pvReserved",        ctypes.c_void_p),
        ("dwReserved",        wintypes.DWORD),
        ("FlagsEx",           wintypes.DWORD),
    ]

#
# SECTION 3: Native Windows Open/Save dialogs (classic style, no tkinter)
#

def windows_file_dialog_open(
    title="Select a file",
    filter_str="All Files\0*.*\0\0",
    initial_dir=None
):
    """
    Opens a classic Windows 'Open File' dialog.
    Returns the chosen path or None if user cancels.
    filter_str example: "CSV Files\0*.csv\0All Files\0*.*\0\0"
    """
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(ofn)

    # Prepare a wide-char buffer for the file name
    file_buffer = (ctypes.c_wchar * MAX_PATH)()
    ofn.lpstrFile = ctypes.cast(file_buffer, wintypes.LPWSTR)
    ofn.nMaxFile = MAX_PATH

    # Filter, e.g. "CSV Files\0*.csv\0All Files\0*.*\0\0"
    ofn.lpstrFilter = filter_str
    ofn.nFilterIndex = 1

    ofn.lpstrTitle = title

    # If you have a starting directory
    if initial_dir:
        ofn.lpstrInitialDir = initial_dir

    # Force user to pick an existing file/path
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST

    success = GetOpenFileNameW(ctypes.byref(ofn))
    if success:
        return file_buffer.value  # The chosen file path
    return None  # user canceled or error


def windows_file_dialog_save(
    title="Save As",
    filter_str="All Files\0*.*\0\0",
    default_filename="",
    initial_dir=None
):
    """
    Opens a classic Windows 'Save File' dialog.
    Returns the chosen file path or None if user cancels.
    """
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(ofn)

    file_buffer = (ctypes.c_wchar * MAX_PATH)()
    # If you have a default filename, set it here
    if default_filename:
        # We can't just assign a Python string to file_buffer;
        # we need to copy it in or do file_buffer.value = default_filename
        # ensuring we don't exceed buffer length.
        file_buffer.value = default_filename

    ofn.lpstrFile = ctypes.cast(file_buffer, wintypes.LPWSTR)
    ofn.nMaxFile = MAX_PATH

    ofn.lpstrFilter = filter_str
    ofn.nFilterIndex = 1
    ofn.lpstrTitle = title

    if initial_dir:
        ofn.lpstrInitialDir = initial_dir

    # Show an overwrite prompt if user tries to overwrite existing file
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT

    success = GetSaveFileNameW(ctypes.byref(ofn))
    if success:
        return file_buffer.value
    return None


#
# SECTION 4: Main Logic — read CSV, exclude Ninja, save
#

def main():
    # 1) Ensure pandas is available
    ensure_pandas_installed()

    import pandas as pd

    # 2) Prompt user to pick an input CSV
    csv_filter = "CSV Files\0*.csv\0All Files\0*.*\0\0"
    input_csv = windows_file_dialog_open(
        title="Select the CSV file to filter",
        filter_str=csv_filter
    )
    if not input_csv:
        print("No file selected. Exiting.")
        sys.exit(0)

    # 3) Try reading the CSV (assuming Windows-likely cp1252 encoding)
    try:
        df = pd.read_csv(input_csv, encoding="cp1252", engine="python")
    except Exception as e:
        print(f"Error reading CSV from {input_csv}:\n{e}")
        sys.exit(1)

    # Validate required columns exist to avoid KeyError later
    required_cols = {"Name", "IP", "HostName"}
    missing = required_cols.difference(df.columns)
    if missing:
        print(f"CSV missing expected column(s): {', '.join(sorted(missing))}")
        sys.exit(1)

    # 4) Exclude any host that has "NinjaRMMAgent"
    #    4a) Find the (IP, HostName) combos with NinjaRMMAgent
    ninja_mask = df['Name'].str.contains("NinjaRMMAgent", na=False)
    hosts_with_ninja = df.loc[ninja_mask, ['IP', 'HostName']].drop_duplicates()

    #    4b) Merge as an anti-join: exclude all rows for those hosts
    merged = df.merge(hosts_with_ninja, on=['IP','HostName'], how='left', indicator=True)
    df_no_ninja = merged[merged['_merge'] == 'left_only'].drop(columns='_merge')

    #    4c) (Optionally) only keep one row per host
    df_no_ninja_unique = df_no_ninja.drop_duplicates(subset=['IP', 'HostName'])

    # 5) Prepare a default output path in the user’s Downloads folder
    current_user = getpass.getuser()
    default_out_folder = f"C:/Users/{current_user}/Downloads"
    default_out_name = "MachinesWithoutNinja.csv"

    # 6) Prompt user to pick the output file
    save_csv = windows_file_dialog_save(
        title="Save filtered CSV",
        filter_str=csv_filter,
        default_filename=default_out_name,
        initial_dir=default_out_folder
    )

    # If user canceled, fallback to the default
    if not save_csv:
        save_csv = os.path.join(default_out_folder, default_out_name)

    # 7) Write out the final filtered CSV
    try:
        df_no_ninja_unique.to_csv(save_csv, index=False)
        print(f"Filtered CSV saved to: {save_csv}")
    except Exception as e:
        print(f"Failed to save CSV: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
