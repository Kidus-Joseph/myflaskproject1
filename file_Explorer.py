#

import subprocess


def open_file_explorer(folder_path):
    subprocess.Popen(['explorer', folder_path])


# Usage
# folder_path = r'C:\path\of\folder'
folder_path = r"C:\Users\kidus\Desktop"
open_file_explorer(folder_path)
