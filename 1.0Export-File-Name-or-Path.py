import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import docx

class FolderExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Export Tool")
        self.root.geometry("600x450")
        
        self.folder_path = tk.StringVar()
        self.export_hidden = tk.BooleanVar()
        self.export_format = tk.StringVar(value="txt")
        self.output_path = tk.StringVar()
        self.output_name = tk.StringVar()
        self.export_option = tk.StringVar(value="both")
        
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        folder_label = ttk.Label(folder_frame, text="Select Folder:")
        folder_label.pack(side=tk.LEFT, padx=5)
        
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path, width=50)
        folder_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        browse_button = ttk.Button(folder_frame, text="Browse", command=self.browse_folder)
        browse_button.pack(side=tk.LEFT, padx=5)
        
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drop)
        
        hidden_check = ttk.Checkbutton(main_frame, text="Include Hidden Files", variable=self.export_hidden)
        hidden_check.pack(anchor=tk.W, pady=5)
        
        option_label = ttk.Label(main_frame, text="Export Options:")
        option_label.pack(anchor=tk.W, pady=5)
        
        option_combo = ttk.Combobox(main_frame, textvariable=self.export_option, values=["names", "paths", "both"])
        option_combo.pack(fill=tk.X, pady=5)
        
        format_label = ttk.Label(main_frame, text="Select Export Format:")
        format_label.pack(anchor=tk.W, pady=5)
        
        format_combo = ttk.Combobox(main_frame, textvariable=self.export_format, values=["txt", "docx", "excel", "csv"])
        format_combo.pack(fill=tk.X, pady=5)
        
        output_label = ttk.Label(main_frame, text="Output Path:")
        output_label.pack(anchor=tk.W, pady=5)
        
        output_entry = ttk.Entry(main_frame, textvariable=self.output_path, width=50)
        output_entry.pack(fill=tk.X, padx=5, pady=5)
        
        output_button = ttk.Button(main_frame, text="Browse", command=self.browse_output_path)
        output_button.pack(anchor=tk.W, pady=5)
        
        name_label = ttk.Label(main_frame, text="Output File Name:")
        name_label.pack(anchor=tk.W, pady=5)
        
        name_entry = ttk.Entry(main_frame, textvariable=self.output_name, width=50)
        name_entry.pack(fill=tk.X, padx=5, pady=5)
        
        export_button = ttk.Button(main_frame, text="Start Export", command=self.start_export)
        export_button.pack(pady=20)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)
        if not self.output_path.get():
            self.output_path.set(folder_selected)
        if not self.output_name.get():
            self.output_name.set(os.path.basename(folder_selected))
    
    def drop(self, event):
        file = event.data.strip('{}')
        self.folder_path.set(file)
        if not self.output_path.get():
            self.output_path.set(file)
        if not self.output_name.get():
            self.output_name.set(os.path.basename(file))
    
    def browse_output_path(self):
        output_selected = filedialog.askdirectory()
        self.output_path.set(output_selected)
    
    def start_export(self):
        folder = self.folder_path.get()
        include_hidden = self.export_hidden.get()
        export_format = self.export_format.get()
        output_path = self.output_path.get()
        output_name = self.output_name.get()
        export_option = self.export_option.get()
        
        if not folder:
            messagebox.showerror("Error", "Please select a folder to export.")
            return
        
        if not output_name:
            output_name = os.path.basename(folder)
        
        output_file = os.path.join(output_path, output_name + '.' + export_format)
        
        file_structure = self.get_file_structure(folder, include_hidden, export_option)
        
        if export_format == 'txt':
            self.export_to_txt(file_structure, output_file)
        elif export_format == 'docx':
            self.export_to_docx(file_structure, output_file)
        elif export_format == 'excel':
            self.export_to_excel(file_structure, output_file)
        elif export_format == 'csv':
            self.export_to_csv(file_structure, output_file)
        
        messagebox.showinfo("Success", f"Files exported successfully to {output_file}")
    
    def get_file_structure(self, folder, include_hidden, export_option):
        file_structure = []
        for root, dirs, files in os.walk(folder):
            root = root.replace(os.sep, '/')
            level = root.replace(folder, '').count('/')
            indent = ' ' * 4 * level
            sub_indent = ' ' * 4 * (level + 1)
            if include_hidden or not os.path.basename(root).startswith('.'):
                if export_option in ['paths', 'both']:
                    file_structure.append(f"{indent}{root}/")
                if export_option in ['names', 'both']:
                    file_structure.append(f"{indent}{os.path.basename(root)}/")
                for f in files:
                    if include_hidden or not f.startswith('.'):
                        if export_option in ['paths', 'both']:
                            file_structure.append(f"{sub_indent}{root}/{f}")
                        if export_option in ['names', 'both']:
                            file_structure.append(f"{sub_indent}{f}")
        return file_structure
    
    def export_to_txt(self, file_structure, output_file):
        with open(output_file, 'w') as f:
            for line in file_structure:
                f.write(line + '\n')
    
    def export_to_docx(self, file_structure, output_file):
        doc = docx.Document()
        for line in file_structure:
            doc.add_paragraph(line)
        doc.save(output_file)
    
    def export_to_excel(self, file_structure, output_file):
        df = pd.DataFrame(file_structure, columns=["File Structure"])
        df.to_excel(output_file, index=False)
    
    def export_to_csv(self, file_structure, output_file):
        df = pd.DataFrame(file_structure, columns=["File Structure"])
        df.to_csv(output_file, index=False)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = FolderExportApp(root)
    root.mainloop()
