import fitz  # PyMuPDF
import os

def pdf_to_jpg(pdf_path, output_dir, resolution=150):
    """
    Convert each page of a PDF to a JPG image.

    :param pdf_path: Path to the PDF file.
    :param output_dir: Directory where JPG images will be saved.
    :param resolution: Resolution for the output images (DPI).
    """
    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    for page_number in range(len(pdf_document)):
        # Select a page
        page = pdf_document[page_number]

        # Render page to a pixmap (image) with specified resolution
        zoom = resolution / 72  # Default resolution is 72 DPI
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        # Save the image as a JPG file
        image_path = os.path.join(output_dir, f"page_{page_number + 1}.jpg")
        pix.save(image_path)
        print(f"Saved: {image_path}")

    # Close the PDF document
    pdf_document.close()
    print("PDF conversion completed.")

# Example usage
if __name__ == "__main__":
    input_pdf = r"C:\Users\alan\Desktop\su7痛车\无标题的页面.pdf"  # Replace with your PDF file path
    output_directory = r"C:\Users\alan\Desktop\su7痛车\线稿_300dpi"  # Replace with your desired output directory

    pdf_to_jpg(input_pdf, output_directory, 300)
