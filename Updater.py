from tkinter import ttk
import tkinter as tk
import requests
import sys
import os
import win32gui
import win32process
import win32api
import threading
import zipfile
import tempfile
import shutil

# Access rights for the process
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010

def RetrieveVersion(classname):
    hwnd = win32gui.FindWindow(classname, None)
    if hwnd:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        access_rights = PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        handle = win32api.OpenProcess(access_rights, False, pid)
        try:
            path = win32process.GetModuleFileNameEx(handle, 0)
            info = win32api.GetFileVersionInfo(path, "\\")
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
            print(f"Program's version: {version}")
            return version
        except Exception as Error:
            print(f"Failed to retrieve version. Reason: {Error}")
            return None
    else:
        print("Failed to find the program.")
        return None

def RetrieveName(classname):
    hwnd = win32gui.FindWindow(classname, None)
    if hwnd:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        acess_rights = PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        handle = win32api.OpenProcess(acess_rights, False, pid)
        try:
            path = win32process.GetModuleFileNameEx(handle, 0)
            name = os.path.basename(path)
            print(f"Program's name: {name}")
            return name
        except Exception as Error:
            print(f"Failed to retrieve name. Reason: {Error}")
            return None
    else:
        print("Failed to find the program.")
        return None

# Owner of the repository where update is coming from (required)
Owner = ""
# Owner's repository where update is coming from (required)
Repository = ""
# Personal Access Token for API calls (optional)
Token = ""
# Path to target directory (required)
TargetDir = os.path.dirname(sys.executable)
# Path to the downloaded ZIP file (leave as is)
ZipPath = None
# Class name of the program (required)
Programclass = ""
# Latest version of the program (required)
CurrentVersion = RetrieveVersion(Programclass)
# Name of the program (required)
Programname = RetrieveName(Programclass)

# Fetch the latest version of your software from Github.
# In this case, I used the Github API to get the latest release with the highest tag name e.g. 1.0.0.0  
def LatestVersion():
    url = f"https://api.github.com/repos/{Owner}/{Repository}/releases"
    headers = {'Authorization': f'token {Token}'}
    response = requests.get(url, headers=headers)
    try:
        if response.status_code == 200:
            data = response.json()
            for release in data:
                if release['tag_name'] >= CurrentVersion:
                    version = release['tag_name']
            print(f"New version found: {version}")
            return version
    except Exception as Error:
        print(f"Failed to get the latest version. Reason: {Error}")
        return None

# Update manager class is responsible for the entire update process.
class UpdateManager:
    def __init__(self, parent):

        # Function to download the update from the Github repository
        def DownloadUpdate(version):
            url = f"https://github.com/{Owner}/{Repository}/releases/download/{version}/{Repository}.zip"
            headers = {'Authorization': f'token {Token}'}
            with requests.get(url, headers=headers, stream=True) as response:
                try:
                    if response.status_code == 200:
                        progressbar['maximum'] = int(response.headers.get('Content-Length', 0))
                        with open('Temp.zip', 'wb') as file:
                            for chunk in response.iter_content(chunk_size=4000):
                                file.write(chunk)
                                progressbar['value'] += 4000
                            zip_path = os.path.join(TargetDir, 'Temp.zip')
                            if zip_path:
                                print("Update downloaded successfully!")
                                return zip_path
                            else:
                                print("Failed to download the update.")
                                return None
                    else:
                        print(f"Failed to download the update. Reason: {response.status_code}")
                        return None
                except Exception as Error:
                    print(f"Error reason: {Error}")
                    return None
                
        # Function to close the program before updating
        def CloseProgram(name):
            try:
                os.system(f"taskkill /f /im {name}")
                print(f"{name} closed successfully!")
                return True
            except Exception as Error:
                print(f"Failed to close the program. Reason: {Error}")
                return False
            
        # Function to open the program after updating
        def OpenProgram(name):
            try:
                os.startfile(name)
                print(f"{name} opened successfully!")
                return True
            except Exception as Error:
                print(f"Failed to open the program. Reason: {Error}")
                return False
            
        # Function to install the update (delete old files and place the new ones in the root dir)
        def InstallUpdate(zip_path, target_dir):
            if CloseProgram(Programname):
                with tempfile.TemporaryDirectory() as temp_dir:
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                    folder = os.listdir(temp_dir)[0]
                    folder_path = os.path.join(temp_dir, folder)

                    for item in os.listdir(target_dir):
                        if item == 'Temp.zip' or item == Programname:
                            os.remove(os.path.join(target_dir, item))

                    for item in os.listdir(folder_path):
                        if item in os.listdir(target_dir):
                            if os.path.isdir(os.path.join(target_dir, item)):
                                shutil.rmtree(os.path.join(target_dir, item))
                                shutil.move(os.path.join(folder_path, item), target_dir)
                            if os.path.isfile(os.path.join(target_dir, item)):
                                os.remove(os.path.join(target_dir, item))
                                shutil.move(os.path.join(folder_path, item), target_dir)
                        else:
                            shutil.move(os.path.join(folder_path, item), target_dir)
            OpenProgram(Programname)

        # Function to download the update               
        def DownloadThread():
            ZipPath = DownloadUpdate(LatestVersion())
            onscreen = "Update downloaded successfully!"
            text.config(text=onscreen)
            InstallUpdate(ZipPath, TargetDir)
        
        # Function to start the download thread
        def DownloadTask():
            threading.Thread(target=DownloadThread).start()

        # The main window for the update process
        Update = tk.Tk()
        Update.title("Updater")
        Update.geometry("300x150")
        x = parent.winfo_x()
        y = parent.winfo_y()
        Update.geometry(f"+{x}+{y}")
        Update.resizable(False, False)
        Update.iconbitmap(default="updater.ico")
        onscreen = "Downloading new version..."
        text = tk.Label(Update, text=onscreen, wraplength=200, justify="center", font=("Arial", 10), pady=20)
        text.pack()
        progressbar = ttk.Progressbar(Update, orient='horizontal', length=200, mode='determinate', value=0, maximum=0)
        progressbar.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        progressbar.pack()
        parent.destroy()
        DownloadTask()
        Update.mainloop()

# Main function that gets called when the program is executed.
def main():
    if LatestVersion() > CurrentVersion:
        global root
        root = tk.Tk()
        root.title("Updater")
        root.geometry("300x150")
        root.resizable(False, False)
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        x = int((sw / 2) - (300 / 2))
        y = int((sh / 2) - (150 / 2))
        root.geometry(f"+{x}+{y}")
        icon = sys._MEIPASS + "/updater.ico"
        root.iconbitmap(default= icon)
        onscreen = "Download new version?"
        text = tk.Label(root, text=onscreen, wraplength=180, justify="center", font=("Arial", 10), pady=20)
        text.pack()
        button_frame = tk.Frame(root)
        button_frame.pack(expand=True, fill='x')
        agree = tk.Button(button_frame, text="Yes", command=lambda: UpdateManager(root))
        agree.pack(side='left', expand=True)
        disagree = tk.Button(button_frame, text="No", command=root.destroy)
        disagree.pack(side='right', expand=True)
        root.mainloop()

if __name__ == "__main__":
    main()