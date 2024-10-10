import os
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
from dotenv import load_dotenv

# Load the .env file
load_dotenv()

# Get the API key from the .env file
api_key = os.getenv('GOOGLE_API_KEY')

if not api_key:
    raise ValueError("GOOGLE_API_KEY environment variable is not set in .env file")

# Configure the genai library with the API key
genai.configure(api_key=api_key)

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
    Create a detailed PowerPoint presentation outline based on the following content:

    {content}

    Generate an outline with the following structure:
    1. Title Slide
    2. Index Slide (will be generated automatically, don't include in your output)
    3-{num_content_slides+2}. Content Slides (5-7 key points per slide)
    {num_content_slides+3}. Conclusion Slide

    For each slide (except the index slide), provide:
    - Slide Title
    - 5-7 Key Points (detailed sentences or ideas)

    Format the output as follows:
    [Slide 1]
    Title: [Presentation Title]

    [Slide 2]
    Title: Index
    - 1: [First Content Slide Title]
    - 2: [Second Content Slide Title]
    - 3: [Third Content Slide Title]
    - ...
    - N: [Last Content Slide Title]
    - Conclusion

    [Slide 3]
    Title: [First Content Slide Title]
    - [Key Point 1 - Detailed sentence]
    - [Key Point 2 - Detailed sentence]
    - [Key Point 3 - Detailed sentence]
    - [Key Point 4 - Detailed sentence]
    - [Key Point 5 - Detailed sentence]
    - [Key Point 6 - Detailed sentence] (optional)
    - [Key Point 7 - Detailed sentence] (optional)

    ... (continue for all content slides and conclusion)

    Ensure each content slide has 5-7 detailed key points.
    Provide comprehensive information while keeping it suitable for a presentation.
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

    for line in output.split('\n'):
        line = line.strip()
        if line.startswith('[Slide'):
            if current_slide:
                slides.append(current_slide)
            current_slide = {'title': '', 'content': []}
        elif line.startswith('Title:'):
            current_slide['title'] = line.split(':', 1)[1].strip()
        elif line.startswith('-'):
            current_slide['content'].append(line)

    if current_slide:
        slides.append(current_slide)

    return slides

def generate_vba_code(slides):
    # Print debugging information
    print("Debugging: Number of slides:", len(slides))
    for i, slide in enumerate(slides):
        print(f"Slide {i}: {slide['title']}")

    vba_code = f"""
Sub CreatePresentation()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim para As TextRange
    
    ' Create a new presentation
    Set ppt = Application.Presentations.Add

    ' Add title slide
    Set sld = ppt.Slides.Add(1, ppLayoutTitle)
    sld.Shapes.Title.TextFrame.TextRange.Text = "{slides[0]['title']}"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)

    ' Add index slide
    Set sld = ppt.Slides.Add(2, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Index"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True

    ' Add index content
"""

    # Add index content
    for i, slide in enumerate(slides[2:], start=1):  # Skip title and index slides
        vba_code += f"""
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{i}: {slide['title'].replace('"', '""')}"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
"""

    # Add Conclusion
    vba_code += """
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conclusion"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
"""

    # Add content slides
    for index, slide in enumerate(slides[2:], start=3):  # Start from slide 3
        vba_code += f"""
    ' Add slide {index}
    Set sld = ppt.Slides.Add({index}, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "{slide['title'].replace('"', '""')}"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText
"""

        for point in slide['content']:
            vba_code += f"""
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{point[1:].strip().replace('"', '""')}"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14
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