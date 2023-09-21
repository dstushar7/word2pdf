import tkinter as tk
from tkinter import filedialog, messagebox
from conversion import word_to_images, images_to_pdf
import os

def convert_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if not file_path:
        return

    # Setting the output folder to be in the Documents/word2pdf directory
    user_documents_path = os.path.expanduser('~/Documents')
    output_folder = os.path.join(user_documents_path, 'word2pdf')

    # Extract the base name of the file (without the path and extension)
    base_name = os.path.basename(file_path)
    file_name_without_extension = os.path.splitext(base_name)[0]
    output_pdf_path = os.path.join(output_folder, f"{file_name_without_extension}.pdf")
    
    try:
        image_paths = word_to_images(file_path, output_folder)
        images_to_pdf(image_paths, output_pdf_path)
        messagebox.showinfo("Success", f"Converted to PDF: {output_pdf_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def run_gui():
    app = tk.Tk()
    app.title("Word to PDF Converter")

    frame = tk.Frame(app, padx=20, pady=20)
    frame.pack(padx=10, pady=10)

    label = tk.Label(frame, text="Select a DOCX file to convert to PDF")
    label.pack(pady=10)

    convert_button = tk.Button(frame, text="Convert DOCX to PDF", command=convert_file)
    convert_button.pack(pady=10)

    app.mainloop()

if __name__ == "__main__":
    run_gui()
