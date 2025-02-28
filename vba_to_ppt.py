import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor

def parse_vba_file(vba_file_path):
    slides_data = []
    current_slide = None
    content_buffer = ""

    with open(vba_file_path, 'r', encoding='utf-8') as file:
        for line_number, line in enumerate(file, 1):
            line = line.strip()
            try:
                if "Set sld = ppt.Slides.Add" in line:
                    if current_slide:
                        current_slide['content'] = content_buffer.strip()
                        slides_data.append(current_slide)
                    current_slide = {'title': '', 'content': ''}
                    content_buffer = ""
                elif ".Shapes.Title.TextFrame.TextRange.Text =" in line:
                    match = re.search(r'"([^"]*)"', line)
                    if match:
                        current_slide['title'] = match.group(1)
                elif ".Text =" in line:
                    match = re.search(r'"([^"]*)"', line)
                    if match:
                        content_buffer += match.group(1)
                elif '& _' in line:
                    match = re.search(r'"([^"]*)"', line)
                    if match:
                        content_buffer += match.group(1) + "\n"
            except ValueError as e:
                print(f"Error processing line {line_number}: {line}")
                print(f"Error details: {str(e)}")

    if current_slide:
        current_slide['content'] = content_buffer.strip()
        slides_data.append(current_slide)

    return slides_data

def create_image_slide(prs, images_dir, page_number):
    """Create a new slide specifically for images from a particular page."""
    # Look for images corresponding to this page
    slide_images = [f for f in os.listdir(images_dir) 
                   if f'page_{page_number}_img_' in f.lower() and 
                   any(f.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif'])]
    
    if not slide_images:
        return None

    # Sort images by their number
    slide_images.sort(key=lambda x: int(re.search(r'img_(\d+)', x).group(1)))
    
    # Create a blank slide for images
    slide_layout = prs.slide_layouts[6]  # Using blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add a title to the image slide
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    title_box.text_frame.text = f"Illustrations from Page {page_number}"
    title_box.text_frame.paragraphs[0].font.size = Pt(28)
    title_box.text_frame.paragraphs[0].font.color.rgb = RGBColor(128, 0, 0)
    
    # Calculate layout for images
    if len(slide_images) == 1:
        # Single image - use larger size
        left_margin = Inches(1)
        top_margin = Inches(1.5)
        img_width = Inches(8)
        img_height = Inches(5.5)
    else:
        # Multiple images - arrange in grid
        left_margin = Inches(0.5)
        top_margin = Inches(1.5)
        img_width = Inches(4.5)
        img_height = Inches(3.5)
        
    for idx, image_file in enumerate(slide_images):
        image_path = os.path.join(images_dir, image_file)
        try:
            # For multiple images, create a 2x2 grid layout
            row = idx // 2
            col = idx % 2
            left = left_margin + (col * (img_width + Inches(0.5)))
            top = top_margin + (row * (img_height + Inches(0.5)))
            
            # Add image to slide
            pic = slide.shapes.add_picture(
                image_path,
                left=left,
                top=top,
                width=img_width,
                height=img_height
            )
            print(f"Successfully added image: {image_file} to image slide for page {page_number}")
        except Exception as e:
            print(f"Error adding image {image_file}: {str(e)}")
    
    return slide

def create_powerpoint(slides_data, output_file):
    # Create a new presentation with a specific theme
    prs = Presentation('themes/Presentation5.pptx')
    
    # Get the images directory
    images_dir = os.path.join('extract', 'images')
    print(f"Looking for images in: {images_dir}")
    
    # Create a list to hold all slides (content and images)
    final_slides = []
    
    for slide_num, slide_data in enumerate(slides_data, 1):
        # Calculate corresponding page number (accounting for title and index slides)
        page_number = slide_num - 2 if slide_num > 2 else None
        
        # Create content slide
        slide_layout = prs.slide_layouts[1]  # Using layout with title and content
        content_slide = prs.slides.add_slide(slide_layout)
        
        title = content_slide.shapes.title
        title.text = slide_data['title']
        print(f"Processing slide {slide_num} with title: {slide_data['title']}")
        
        # Set the title color to maroon
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(128, 0, 0)
        
        content = content_slide.placeholders[1]
        tf = content.text_frame
        tf.clear()  # Clear existing content
        
        paragraphs = slide_data['content'].split('vbNewLine')
        for para in paragraphs:
            p = tf.add_paragraph()
            p.text = para.strip()
            p.level = 0
            p.font.size = Pt(18)
        
        # After each content slide (except title and index), check for and add image slide
        if page_number and page_number > 0:
            image_slide = create_image_slide(prs, images_dir, page_number)
            if image_slide:
                print(f"Added image slide after content slide {slide_num}")

    # Remove empty placeholder shapes
    for slide in prs.slides:
        shapes_to_remove = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and not shape.has_text_frame:
                shapes_to_remove.append(shape)
        for shape in shapes_to_remove:
            sp = shape._element
            sp.getparent().remove(sp)

    prs.save(output_file)

def main():
    vba_file_path = 'create_presentation.vba'
    output_file = os.path.join('output', 'generated_presentation.pptx')
    
    slides_data = parse_vba_file(vba_file_path)
    create_powerpoint(slides_data, output_file)
    print(f"PowerPoint presentation created: {output_file}")
    return output_file

if __name__ == "__main__":
    main()
