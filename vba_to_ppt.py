import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

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

def create_powerpoint(slides_data, output_file):
    prs = Presentation()
    
    for slide_data in slides_data:
        slide_layout = prs.slide_layouts[1]  # Using layout with title and content
        slide = prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = slide_data['title']
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()  # Clear existing content
        
        paragraphs = slide_data['content'].split('vbNewLine')
        for para in paragraphs:
            p = tf.add_paragraph()
            p.text = para.strip()
            p.level = 0
            p.font.size = Pt(18)

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
