import pdfplumber
import os
import fitz  # PyMuPDF
import re

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

    # Dictionary to store image titles
    image_titles = {}
    images_found = False

    # Get the base name of the PDF file (without extension)
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

    # First pass: Extract text and identify potential image titles
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                tables = page.extract_tables()

                # Save text content
                output_file = os.path.join(output_dir, f"{pdf_name}_page_{page_num}.txt")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(text + "\n\n")
                    for table in tables:
                        for row in table:
                            f.write(" | ".join(str(cell) for cell in row) + "\n")
                        f.write("\n")

                # Look for potential image titles in the text
                # Common patterns for image titles
                title_patterns = [
                    r'(?:Figure|Fig\.|FIGURE)\s*(\d+)[:\.]?\s*([^\n\.]+)',
                    r'(?:Diagram|DIAGRAM)\s*(\d+)[:\.]?\s*([^\n\.]+)',
                    r'(?:Image|IMAGE)\s*(\d+)[:\.]?\s*([^\n\.]+)',
                    r'(?:Illustration|ILLUSTRATION)\s*(\d+)[:\.]?\s*([^\n\.]+)'
                ]

                for pattern in title_patterns:
                    matches = re.finditer(pattern, text, re.IGNORECASE)
                    for match in matches:
                        fig_num = match.group(1)
                        title = match.group(2).strip()
                        image_titles[f"page_{page_num}_img_{fig_num}"] = title

        # Second pass: Extract and save images with titles
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            image_list = page.get_images()
            
            if image_list:
                images_found = True
                
            for img_index, img in enumerate(image_list, 1):
                xref = img[0]
                try:
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # Generate image filename
                    image_filename = f"{pdf_name}_page_{page_num + 1}_img_{img_index}.{image_ext}"
                    image_path = os.path.join(images_dir, image_filename)
                    
                    # Save image
                    with open(image_path, "wb") as image_file:
                        image_file.write(image_bytes)

                    # Save image title if found, otherwise use a default title
                    title_key = f"page_{page_num + 1}_img_{img_index}"
                    if title_key not in image_titles:
                        # Try to extract title from surrounding text
                        surrounding_text = page.get_text("text", clip=(img[1], img[2], img[3], img[4]))
                        # Look for potential title in the text above the image
                        lines = surrounding_text.split('\n')
                        for line in lines:
                            if line.strip() and len(line.strip()) > 5:  # Reasonable title length
                                image_titles[title_key] = line.strip()
                                break
                        if title_key not in image_titles:
                            image_titles[title_key] = f"Figure {img_index}"
                except Exception as img_error:
                    print(f"Error extracting image {img_index} on page {page_num + 1}: {img_error}")
                    continue

        # Save image titles to a file
        titles_file = os.path.join(output_dir, "image_titles.txt")
        with open(titles_file, 'w', encoding='utf-8') as f:
            for key, title in image_titles.items():
                f.write(f"{key}|{title}\n")

        if not images_found:
            print(f"No images found in {pdf_path}")
            # Create a marker file to indicate no images were found
            with open(os.path.join(images_dir, "no_images.txt"), 'w') as f:
                f.write("No images were found in this PDF.")

        return images_dir, image_titles

    except Exception as e:
        print(f"Error processing PDF: {e}")
        # Ensure the images directory exists even if there was an error
        os.makedirs(images_dir, exist_ok=True)
        # Create a marker file to indicate an error occurred
        with open(os.path.join(images_dir, "extraction_error.txt"), 'w') as f:
            f.write(f"Error processing PDF: {str(e)}")
        return images_dir, {}

# Remove the example usage from here
