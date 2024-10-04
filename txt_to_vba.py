import os
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold


GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY")
if not GOOGLE_API_KEY:
    raise ValueError("GOOGLE_API_KEY environment variable is not set")
genai.configure(api_key=GOOGLE_API_KEY)

def read_input_files(folder_path):
    combined_content = ""
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                combined_content += file.read() + "\n\n"
    return combined_content[:8000] 

def generate_outline_with_gemini(content, num_content_slides):
    model = genai.GenerativeModel('gemini-pro')
    
    prompt = f"""
    Create a PowerPoint presentation outline based on the following content:

    {content}

    Generate an outline with the following structure:
    1. Title Slide
    2-{num_content_slides+1}. Content Slides (varying number of main points and details per slide)
    {num_content_slides+2}. Conclusion Slide

    For each slide, provide:
    - Slide Title
    - Multiple Main Points (2-5 bullet points)
    - Varying number of details or sub-points for each main point

    Format the output as follows:
    [Slide 1]
    Title: [Slide Title]
    - [Main Point 1]
      • [Detail 1]
      • [Detail 2]
    - [Main Point 2]
      • [Detail 1]
    - [Main Point 3]
      • [Detail 1]
      • [Detail 2]
      • [Detail 3]

    [Slide 2]
    Title: [Slide Title]
    - [Main Point 1]
    - [Main Point 2]
      • [Detail 1]
    - [Main Point 3]
      • [Detail 1]
      • [Detail 2]
    - [Main Point 4]

    ... (continue for all slides)

    Ensure each content slide has a varying number of main points (2-5) with a different number of details for each point.
    Keep the content concise and suitable for a presentation.
    """

    generation_config = {
        "temperature": 0.7,
        "top_p": 1,
        "top_k": 1,
        "max_output_tokens": 2048,
    }

    safety_settings = [
        {
            "category": HarmCategory.HARM_CATEGORY_HARASSMENT,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
        {
            "category": HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
        {
            "category": HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
        {
            "category": HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
    ]

    response = model.generate_content(
        prompt,
        generation_config=generation_config,
        safety_settings=safety_settings
    )

    return response.text

def parse_gemini_output(output):
    slides = []
    current_slide = None
    current_point = None

    for line in output.split('\n'):
        line = line.strip()
        if line.startswith('[Slide'):
            if current_slide:
                slides.append(current_slide)
            current_slide = {'title': '', 'content': []}
        elif line.startswith('Title:'):
            current_slide['title'] = line.split(':', 1)[1].strip()
        elif line.startswith('-'):
            current_point = line
            current_slide['content'].append(current_point)
        elif line.startswith('•'):
            if current_point:
                current_slide['content'].append(f"  {line}")

    if current_slide:
        slides.append(current_slide)

    return slides

def generate_vba_code(slides):
    vba_code = """
Sub CreatePresentation()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim para As TextRange
    
    ' Create a new presentation
    Set ppt = Application.Presentations.Add
    """

    for index, slide in enumerate(slides, start=1):
        vba_code += f"""
    ' Add slide {index}
    Set sld = ppt.Slides.Add({index}, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "{slide['title']}"
    Set shp = sld.Shapes(2)
    Set tf = shp.TextFrame
    tf.DeleteText
    """

        for point in slide['content']:
            if point.startswith('-'):
                vba_code += f"""
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{point[1:].strip()}"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    """
            elif point.startswith('  •'):
                vba_code += f"""
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{point[3:].strip()}"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 0.8
    para.IndentLevel = 2
    """

    vba_code += """
End Sub
"""

    return vba_code

def main():
    input_folder = 'extract'  # Name of your input folder containing txt files
    output_file = 'create_presentation.vba'
    num_content_slides = 6  # Customize this number as needed

    if not os.path.exists(input_folder):
        print(f"Error: Input folder '{input_folder}' not found.")
        return

    content = read_input_files(input_folder)
    if not content:
        print("No text files found in the input folder.")
        return

    gemini_output = generate_outline_with_gemini(content, num_content_slides)
    slides = parse_gemini_output(gemini_output)
    vba_code = generate_vba_code(slides)
    
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(vba_code)

    print(f"VBA code generated and saved to '{output_file}'")

if __name__ == "__main__":
    main()