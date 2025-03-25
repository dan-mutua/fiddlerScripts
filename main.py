import subprocess
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def update_zap_path(zap_path):
    """Update the ZAP file path in index.py"""
    index_path = os.path.join(os.path.dirname(__file__), 'index.py')
    
    with open(index_path, 'r') as file:
        content = file.read()
        
    updated_content = content.replace(
        'ZAP_FILE_PATH = r"..."', 
        f'ZAP_FILE_PATH = r"{zap_path}"'
    )
    
    with open(index_path, 'w') as file:
        file.write(updated_content)

def run_script(script_name, file_path=None):
    """Run a Python script and return success status"""
    print("wait for magic to happen......")
    print(f"\nRunning {script_name}...")
    cmd = [sys.executable, script_name]
    if file_path:
        cmd.append(file_path)
    result = subprocess.run(
        cmd,
        cwd=os.path.dirname(os.path.abspath(__file__))
    )
    return result.returncode == 0

def main():
    root = tk.Tk()
    root.title("ZAP Analysis Tool")
    root.geometry("400x200")
    
    label = tk.Label(root, text="Please select your .zap file to begin analysis:")
    label.pack(pady=20)
    
    def select_file():
        file_path = filedialog.askopenfilename(
            title="Select ZAP File",
            filetypes=[
                ("ZAP files", "*.zap"),
                ("All files", "*.*")  
            ],
            defaultextension=".zap"
        )
        if file_path:
            root.destroy()
            try:
                update_zap_path(file_path)
                
                scripts = [
                    'index.py',
                    'searchMsDocs.py',
                    'ariSearch.py'
                ]
                
                for script in scripts:
                    if not run_script(script, file_path):
                        messagebox.showerror(
                            "Error",
                            f"Failed to run {script}"
                        )
                        sys.exit(1)
                
                messagebox.showinfo(
                    "Success",
                    "Analysis completed successfully!\nResults saved to Excel file."
                )
                
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"An error occurred: {str(e)}"
                )
                sys.exit(1)
    
    select_button = tk.Button(
        root,
        text="Select ZAP File",
        command=select_file
    )
    select_button.pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    main()