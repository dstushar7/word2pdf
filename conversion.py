from docx2pdf import convert
from pdf2image import convert_from_path
from PIL import Image
from utils import create_directory

def word_to_images(word_file_path, output_image_folder):
    create_directory(output_image_folder)

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

    os.remove(pdf_file_path)
    return image_paths

def images_to_pdf(image_paths, output_pdf_path):
    img = Image.open(image_paths[0])
    img.save(output_pdf_path, "PDF", resolution=100.0, save_all=True, append_images=[Image.open(image_path) for image_path in image_paths[1:]])
