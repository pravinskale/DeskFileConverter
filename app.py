import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pdftoexcel import PDFTOExcelConverter
upload_folder = "uploads"
def browse_file():
    file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("PDF files", "*.pdf")] )
    if file_path:
        file_path_var.set(file_path)

def convert_file():
    file_path = file_path_var.get()
    if file_path:
        convert_pdf_to_excel(file_path)
        #messagebox.showinfo("Convert", f"Converting file:\n{file_path}")
    else:
        messagebox.showwarning("No file", "Please select a file first.")
def convert_pdf_to_excel(file_path):
    try:
        ValidateFile(file_path)
        converter = PDFTOExcelConverter()
        excel_file_path = file_path.replace('.pdf', '.xlsx')
        # Call pdf_to_excel method
        converter.pdf_to_excel(
            pdfToConvert=file_path,
            convertedExcel= excel_file_path
        )
        messagebox.showinfo("Success", f"File converted successfully at {excel_file_path}!")
    except PermissionError as e:
        messagebox.showerror("Permission Error", f"Check if file is accessed/opened by other process:\n{e}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert file:\n{e}")
      
def ValidateFile(file, type='.pdf'):
    """
    Validate the uploaded file.
    """
    if not file:
        #throw an exception if no file is provided
        raise Exception("No file provided")
    if not file.endswith(type):
        raise Exception(f"File is not a {type} file")
    
# Main window
root = tk.Tk()
root.title("🛠️ File Converter")
root.geometry("500x200")
#root.res()


# Style
style = ttk.Style()
style.configure("TButton", padding=6, font=("Segoe UI", 10))
style.configure("TLabel", font=("Segoe UI", 10))

# File selection frame
frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

file_path_var = tk.StringVar()

ttk.Label(frame, text="Selected File:").grid(row=0, column=0, sticky="w")
file_entry = ttk.Entry(frame, textvariable=file_path_var, width=50, state="readonly")
file_entry.grid(row=1, column=0, columnspan=2, pady=5)

browse_btn = ttk.Button(frame, text="📁 Browse", command=browse_file)
browse_btn.grid(row=2, column=0, pady=10, sticky="w")

convert_btn = ttk.Button(frame, text="🔄 Convert", command=convert_file)
convert_btn.grid(row=2, column=1, pady=10, sticky="e")

# Run the app
root.mainloop()
