import pdfplumber
import os
import fitz  # PyMuPDF
from PIL import Image
import io

def extract_text_from_pdf(pdf_path, output_dir):
    """
    Maintain backward compatibility with existing code
    """
    return extract_text_and_images_from_pdf(pdf_path, output_dir)

def extract_text_and_images_from_pdf(pdf_path, output_dir):
    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Create a subdirectory for images
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    # Get the base name of the PDF file (without extension)
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

    # Extract text and tables using pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                tables = page.extract_tables()

                output_file = os.path.join(output_dir, f"{pdf_name}_page_{page_num}.txt")
                
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(text + "\n\n")
                    for table in tables:
                        for row in table:
                            f.write(" | ".join(str(cell) for cell in row) + "\n")
                        f.write("\n")

                print(f"Text and data extracted from page {page_num} and saved in {output_file}")

        # Extract images using PyMuPDF
        doc = fitz.open(pdf_path)
        image_count = 0
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            image_list = page.get_images()
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # Save image
                image_filename = f"{pdf_name}_page_{page_num + 1}_img_{img_index + 1}.{image_ext}"
                image_path = os.path.join(images_dir, image_filename)
                
                with open(image_path, "wb") as image_file:
                    image_file.write(image_bytes)
                image_count += 1
                
                print(f"Extracted image {image_filename}")
        
        print(f"Extraction complete. Found {image_count} images.")
        return images_dir  # Return the path where images are stored

    except Exception as e:
        print(f"Error processing PDF: {e}")
        return None

# Remove the example usage from here
