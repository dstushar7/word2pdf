import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
from pdf2image import convert_from_path
from PIL import Image


def word_to_images(word_file_path, output_image_folder):
    # Check if the output folder exists
    if not os.path.exists(output_image_folder):
        os.makedirs(output_image_folder)

    # Convert Word to PDF
    pdf_file_path = os.path.join(output_image_folder, "temp.pdf")
    convert(word_file_path, pdf_file_path)

    # Convert PDF to images
    images = convert_from_path(pdf_file_path)

    # Save images
    image_paths = []
    for i, image in enumerate(images):
        image_path = os.path.join(output_image_folder, f"page_{i + 1}.png")
        image.save(image_path, 'PNG')
        image_paths.append(image_path)

    # Cleanup temporary PDF
    os.remove(pdf_file_path)

    return image_paths


def images_to_pdf(image_paths, output_pdf_path):
    # Use the PIL library to combine images into a PDF
    img = Image.open(image_paths[0])
    img.save(output_pdf_path, "PDF", resolution=100.0, save_all=True, append_images=[Image.open(image_path) for image_path in image_paths[1:]])


def convert_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if not file_path:
        return

    # Setting the output folder to be in the Documents/word2pdf directory
    user_documents_path = os.path.expanduser('~/Documents')
    output_folder = os.path.join(user_documents_path, 'word2pdf')
    
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Extract the base name of the file (without the path and extension)
    base_name = os.path.basename(file_path)  # e.g., "document.docx"
    file_name_without_extension = os.path.splitext(base_name)[0]  # e.g., "document"
    output_pdf_path = os.path.join(output_folder, f"{file_name_without_extension}.pdf")  # e.g., "Documents/word2pdf/document.pdf"
    
    try:
        image_paths = word_to_images(file_path, output_folder)
        images_to_pdf(image_paths, output_pdf_path)
        messagebox.showinfo("Success", f"Converted to PDF: {output_pdf_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")



app = tk.Tk()
app.title("Word to PDF Converter")

frame = tk.Frame(app, padx=20, pady=20)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="Select a DOCX file to convert to PDF")
label.pack(pady=10)

convert_button = tk.Button(frame, text="Convert DOCX to PDF", command=convert_file)
convert_button.pack(pady=10)

app.mainloop()
