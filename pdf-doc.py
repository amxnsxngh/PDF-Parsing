import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pdf2docx import Converter


def convert_pdf_to_docx(pdf_path, docx_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()
        messagebox.showinfo("Success", f"Conversion complete. File saved as: {docx_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        # Prompt the user to select a destination folder
        folder_path = filedialog.askdirectory()
        if folder_path:
            # Create the DOCX file path
            docx_file = f"{folder_path}/{pdf_path.rsplit('/', 1)[-1].rsplit('.', 1)[0]}.docx"
            convert_pdf_to_docx(pdf_path, docx_file)

# Create the main window
root = tk.Tk()
root.title("PDF to DOCX Converter")

# Create and place the browse button
browse_button = tk.Button(root, text="Select PDF and Convert", command=browse_pdf)
browse_button.pack(pady=20)

# Run the application
root.mainloop()
