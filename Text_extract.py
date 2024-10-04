import pdfplumber
import os

def extract_text_from_pdf(pdf_path, output_dir):
    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Get the base name of the PDF file (without extension)
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

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

        print(f"Text extraction complete for {pdf_path}")
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")

# Remove the example usage from here
