from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def pdf_to_word(input_file, output_file):
    try:
        # Convert the .pdf file to .docx
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
        print(f"Conversion successful! Word file saved as {output_file}")
        messagebox.showinfo("Success", f"Word file saved as {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

def select_file():
    # Create a file dialog for selecting the .pdf file
    input_file = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )

    if input_file:
        output_file = os.path.splitext(input_file)[0] + '.docx'
        pdf_to_word(input_file, output_file)
    else:
        print("No file selected")
        messagebox.showwarning("No Selection", "No file selected")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    select_file()
