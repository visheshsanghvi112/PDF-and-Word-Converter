from docx2pdf import convert
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def word_to_pdf(input_file, output_file):
    try:
        # Convert the .docx file to .pdf
        convert(input_file, output_file)
        print(f"Conversion successful! PDF saved as {output_file}")
        messagebox.showinfo("Success", f"PDF saved as {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

def select_file():
    # Create a file dialog for selecting the .docx file
    input_file = filedialog.askopenfilename(
        title="Select a Word file",
        filetypes=[("Word files", "*.docx")]
    )

    if input_file:
        output_file = os.path.splitext(input_file)[0] + '.pdf'
        word_to_pdf(input_file, output_file)
    else:
        print("No file selected")
        messagebox.showwarning("No Selection", "No file selected")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    select_file()
