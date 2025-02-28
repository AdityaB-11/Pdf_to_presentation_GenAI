import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor

# Define themes with specific files
THEMES = {
    'theme1': {
        'name': 'Theme 1',
        'file': 'Presentation1.pptx'
    },
    'theme2': {
        'name': 'Theme 2',
        'file': 'Presentation2.pptx'
    },
    'theme3': {
        'name': 'Theme 3',
        'file': 'Presentation3.pptx'
    },
    'theme4': {
        'name': 'Theme 4',
        'file': 'Presentation4.pptx'
    }
}

def list_available_themes():
    """Return a list of available themes."""
    return [{'key': key, 'name': theme['name']} for key, theme in THEMES.items()]

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

def apply_theme_to_slide(slide, theme):
    """Apply theme colors and styles to a slide."""
    # Apply theme to title
    if slide.shapes.title:
        title_frame = slide.shapes.title.text_frame
        title_frame.paragraphs[0].font.color.rgb = theme['title_color']
        title_frame.paragraphs[0].font.size = theme['font_size']['title']
    
    # Apply theme to body
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                if shape != slide.shapes.title:  # Skip title
                    paragraph.font.color.rgb = theme['body_color']
                    paragraph.font.size = theme['font_size']['body']

def get_image_titles():
    """Read image titles from the file."""
    titles_file = os.path.join('extract', 'image_titles.txt')
    titles = {}
    if os.path.exists(titles_file):
        with open(titles_file, 'r', encoding='utf-8') as f:
            for line in f:
                key, title = line.strip().split('|', 1)
                titles[key] = title
    return titles

def create_image_slide(prs, images_dir, page_number):
    """Create a new slide specifically for images."""
    slide_images = [f for f in os.listdir(images_dir) 
                   if f'page_{page_number}_img_' in f.lower() and 
                   any(f.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif'])]
    
    if not slide_images:
        return None

    slide_images.sort(key=lambda x: int(re.search(r'img_(\d+)', x).group(1)))
    
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = f"Figures from Section {page_number}"
    
    # Calculate layout for images
    if len(slide_images) == 1:
        left_margin = Inches(1)
        top_margin = Inches(1.5)
        img_width = Inches(8)
        img_height = Inches(5.5)
    else:
        left_margin = Inches(0.5)
        top_margin = Inches(1.5)
        img_width = Inches(4.5)
        img_height = Inches(3.5)
        
    for idx, image_file in enumerate(slide_images):
        image_path = os.path.join(images_dir, image_file)
        try:
            row = idx // 2
            col = idx % 2
            left = left_margin + (col * (img_width + Inches(0.5)))
            top = top_margin + (row * (img_height + Inches(0.5)))
            
            slide.shapes.add_picture(
                image_path,
                left=left,
                top=top,
                width=img_width,
                height=img_height
            )
            print(f"Added image: {image_file}")
        except Exception as e:
            print(f"Error adding image {image_file}: {str(e)}")
    
    return slide

def create_powerpoint(slides_data, output_file, theme_key='theme1'):
    """Create PowerPoint with specified theme."""
    print(f"Creating PowerPoint with theme: {theme_key}")
    
    # Get theme configuration
    theme = THEMES.get(theme_key, THEMES['theme1'])
    theme_path = os.path.join('themes', theme['file'])
    
    # Create presentation with selected theme
    prs = Presentation(theme_path)
    
    # Get the images directory
    images_dir = os.path.join('extract', 'images')
    print(f"Looking for images in: {images_dir}")
    
    for slide_num, slide_data in enumerate(slides_data, 1):
        # Calculate corresponding page number
        page_number = slide_num - 2 if slide_num > 2 else None
        
        # Create content slide
        slide_layout = prs.slide_layouts[1]  # Using layout with title and content
        content_slide = prs.slides.add_slide(slide_layout)
        
        # Set slide title
        title = content_slide.shapes.title
        title.text = slide_data['title']
        print(f"Processing slide {slide_num} with title: {slide_data['title']}")
        
        # Add content
        content = content_slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        paragraphs = slide_data['content'].split('vbNewLine')
        for para in paragraphs:
            p = tf.add_paragraph()
            p.text = para.strip()
            p.level = 0
        
        # Add image slide if needed
        if page_number and page_number > 0:
            image_slide = create_image_slide(prs, images_dir, page_number)
            if image_slide:
                print(f"Added image slide after content slide {slide_num}")

    prs.save(output_file)
    print(f"Presentation saved with theme {theme_key} to: {output_file}")
    return output_file

def main():
    vba_file_path = 'create_presentation.vba'
    output_file = os.path.join('output', 'generated_presentation.pptx')
    
    # List available themes
    print("\nAvailable themes:")
    for theme in list_available_themes():
        print(f"- {theme['name']} (key: {theme['key']})")
    
    slides_data = parse_vba_file(vba_file_path)
    ppt_path = create_powerpoint(slides_data, output_file)
    return ppt_path

if __name__ == "__main__":
    main()
