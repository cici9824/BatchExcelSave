# Internal filename: 'E:\\Software\\OPPO\\OPPO_Multilingual\\!Tools\\Save-Excel\\v4\\Clearly_Local_Batch_Excel_SaveAs_v3.py'
# Bytecode version: 3.13.0rc3 (3571)
# Source timestamp: 1970-01-01 00:00:00 UTC (0)

import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pythoncom
import win32com.client as win32
def find_excel_files(folder):
    """Recursively find all Excel files (.xls, .xlsx, .xlsm)"""
    exts = {'.xlsm', '.xlsx', '.xls'}
    files = []
    for root, dirs, filenames in os.walk(folder):
        for f in filenames:
            if os.path.splitext(f)[1].lower() in exts:
                files.append(os.path.join(root, f))
    return files
def save_excel_file(src_path, dst_path):
    """Open Excel file with COM and save as xlsx format"""
    try:
        pythoncom.CoInitialize()
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(src_path)
        wb.SaveAs(dst_path, FileFormat=51)
        wb.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()
        return (True, f'✅ Saved successfully: {dst_path}')
    except AttributeError as e:
        return (False, f'❌ Failed to save: {src_path}\nError (AttributeError): {e}')
    except Exception as e:
        return (False, f'❌ Failed to save: {src_path}\nError:\n{e}')
def browse_src_folder():
    folder = filedialog.askdirectory()
    if folder:
        src_entry.delete(0, tk.END)
        src_entry.insert(0, folder)
def browse_dst_folder():
    folder = filedialog.askdirectory()
    if folder:
        dst_entry.delete(0, tk.END)
        dst_entry.insert(0, folder)
def start_processing():
    src_folder = src_entry.get()
    dst_folder = dst_entry.get()
    if not os.path.isdir(src_folder):
        messagebox.showerror('Error', 'Please select a valid source folder')
        return
    else:
        if not os.path.isdir(dst_folder):
            messagebox.showerror('Error', 'Please select a valid destination folder')
            return
        else:
            files = find_excel_files(src_folder)
            if not files:
                messagebox.showinfo('Info', 'No Excel files found in the source folder')
                return
            else:
                log_text.delete(1.0, tk.END)
                root.update()
                for file_path in files:
                    rel_path = os.path.relpath(file_path, src_folder)
                    rel_path_no_ext = os.path.splitext(rel_path)[0]
                    dst_path = os.path.join(dst_folder, rel_path_no_ext + '.xlsx')
                    dst_path = os.path.normpath(dst_path)
                    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                    success, msg = save_excel_file(file_path, dst_path)
                    log_text.insert(tk.END, msg + '\n')
                    log_text.see(tk.END)
                    root.update()
                    time.sleep(0.3)
                messagebox.showinfo('Done', 'All files processed successfully')
root = tk.Tk()
root.title('Clearly Local Batch Excel SaveAs Tool')
root.geometry('700x500')
tk.Label(root, text='Source folder (containing Excel files):').pack(anchor='w', padx=10, pady=5)
frame_src = tk.Frame(root)
frame_src.pack(fill='x', padx=10)
src_entry = tk.Entry(frame_src)
src_entry.pack(side='left', fill='x', expand=True)
btn_browse_src = tk.Button(frame_src, text='Browse', command=browse_src_folder)
btn_browse_src.pack(side='left', padx=5)
tk.Label(root, text='Destination folder (for saved files):').pack(anchor='w', padx=10, pady=5)
frame_dst = tk.Frame(root)
frame_dst.pack(fill='x', padx=10)
dst_entry = tk.Entry(frame_dst)
dst_entry.pack(side='left', fill='x', expand=True)
btn_browse_dst = tk.Button(frame_dst, text='Browse', command=browse_dst_folder)
btn_browse_dst.pack(side='left', padx=5)
btn_start = tk.Button(root, text='Start Processing', command=start_processing)
btn_start.pack(pady=15)
tk.Label(root, text='Log output:').pack(anchor='w', padx=10)
log_text = scrolledtext.ScrolledText(root, height=15)
log_text.pack(fill='both', padx=10, pady=5, expand=True)
root.mainloop()
