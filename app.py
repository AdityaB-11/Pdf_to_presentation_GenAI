import os
import traceback
import shutil
import requests
import uuid
import urllib.parse
from io import BytesIO
from flask import Flask, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename
import Text_extract
import txt_to_vba
import vba_to_ppt

app = Flask(__name__, static_folder='static')

UPLOAD_FOLDER = 'uploads'
EXTRACT_FOLDER = 'extract'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXTRACT_FOLDER'] = EXTRACT_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Define themes with preview images
THEMES = {
    'theme1': {
        'name': 'Theme 1',
        'file': 'Presentation1.pptx',
        'preview': 'presentation1.png'
    },
    'theme2': {
        'name': 'Theme 2',
        'file': 'Presentation2.pptx',
        'preview': 'presentation2.png'
    },
    'theme3': {
        'name': 'Theme 3',
        'file': 'Presentation3.pptx',
        'preview': 'presentation3.png'
    },
    'theme4': {
        'name': 'Theme 4',
        'file': 'Presentation4.pptx',
        'preview': 'presentation4.png'
    }
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_folders():
    """Clean up the extract and uploads folders"""
    try:
        # Clean extract folder
        extract_folder = app.config['EXTRACT_FOLDER']
        for item in os.listdir(extract_folder):
            item_path = os.path.join(extract_folder, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):  # For subdirectories like 'images'
                shutil.rmtree(item_path)
        print("Cleaned extract folder")

        # Clean uploads folder
        uploads_folder = app.config['UPLOAD_FOLDER']
        for file in os.listdir(uploads_folder):
            file_path = os.path.join(uploads_folder, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print("Cleaned uploads folder")

    except Exception as e:
        print(f"Error during cleanup: {str(e)}")

def fetch_images_for_topic(topic, slide_titles, num_images=4):
    """Fetch relevant images for the topic and slides from web search"""
    try:
        print(f"Fetching images for topic: {topic}")
        images_dir = os.path.join(app.config['EXTRACT_FOLDER'], 'images')
        os.makedirs(images_dir, exist_ok=True)
        
        # Create image_titles.txt file
        titles_file = os.path.join(app.config['EXTRACT_FOLDER'], 'image_titles.txt')
        
        # List to keep track of saved images
        saved_images = []
        image_titles = {}
        
        # Search terms include the main topic and selected slide titles
        search_terms = [topic]
        if slide_titles:
            # Add a few slide titles as search terms (skip title and agenda slides)
            content_titles = [title for title in slide_titles if topic.lower() not in title.lower() 
                            and "agenda" not in title.lower() 
                            and "overview" not in title.lower()
                            and "conclusion" not in title.lower()][:3]
            search_terms.extend(content_titles)
        
        # Prepare the Unsplash API endpoint
        # Note: In production, you would use a proper API key
        # Here we're using a simple and limited approach
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        for idx, search_term in enumerate(search_terms):
            if len(saved_images) >= num_images:
                break
                
            # Try to get a relevant image
            try:
                # Encode the search term for URL
                encoded_term = urllib.parse.quote(search_term)
                
                # Try Unsplash source (a simple API that doesn't require authentication)
                # We're using different sizes to get variety
                if idx == 0:  # Main topic - larger image
                    url = f"https://source.unsplash.com/1200x800/?{encoded_term}"
                else:
                    url = f"https://source.unsplash.com/800x600/?{encoded_term}"
                
                response = requests.get(url, headers=headers, timeout=5)
                
                if response.status_code == 200:
                    # Generate a unique filename
                    img_id = str(uuid.uuid4())[:8]
                    page_num = idx + 1  # Start with page 1
                    img_num = len(saved_images) + 1
                    
                    # Define filename format similar to PDF extraction
                    file_extension = '.jpg'  # Unsplash typically returns JPGs
                    image_filename = f"topic_{page_num}_img_{img_num}{file_extension}"
                    image_path = os.path.join(images_dir, image_filename)
                    
                    # Save the image
                    with open(image_path, 'wb') as f:
                        f.write(response.content)
                    
                    # Create a key for the image titles - this is critical for matching in vba_to_ppt.py
                    # The key needs to match the pattern in create_image_slide: page_{page_number}_img_{img_index}
                    img_key = f"page_{page_num}_img_{img_num}"
                    
                    # Set the image title
                    image_titles[img_key] = f"{search_term}"
                    saved_images.append(image_path)
                    
                    print(f"Saved image: {image_filename} with title: {search_term}, key: {img_key}")
            
            except Exception as img_error:
                print(f"Error fetching image for '{search_term}': {str(img_error)}")
                continue
        
        # Write image titles to file
        with open(titles_file, 'w', encoding='utf-8') as f:
            for key, title in image_titles.items():
                f.write(f"{key}|{title}\n")
        
        print(f"Saved {len(saved_images)} images for the presentation")
        return saved_images
    
    except Exception as e:
        print(f"Error in fetch_images_for_topic: {str(e)}")
        return []

def generate_topic_content(topic, details="", slide_count=8, presentation_rules=""):
    """Generate presentation content based on a topic using Gemini"""
    # Format the prompt for Gemini
    prompt = f"""
    Create a comprehensive presentation outline on the topic: {topic}.
    
    Additional details: {details}
    
    For this outline, include:
    1. An engaging title slide
    2. A brief agenda/overview slide
    3. {slide_count} detailed content slides with key points
    4. A conclusion slide with summary and takeaways
    """
    
    # Add presentation rules if provided
    if presentation_rules:
        prompt += f"""
    
    PRESENTATION RULES:
    {presentation_rules}
    """
    
    prompt += """
    Format your response using the following structure:
    
    TITLE: [Presentation Title]
    SUBTITLE: [Optional Subtitle]
    
    SLIDE 1: [Slide Title]
    - [Bullet point 1]
    - [Bullet point 2]
    
    SLIDE 2: [Slide Title]
    - [Bullet point 1]
    - [Bullet point 2]
    
    [... and so on for each slide]
    """
    
    # Call Gemini to generate the outline
    gemini_output = txt_to_vba.generate_outline_with_gemini(prompt, slide_count)
    return gemini_output

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'
        }
        for key, theme in THEMES.items()
    ]
    
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', message='No file part', themes=themes)
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', message='No selected file', themes=themes)
        
        # Get selected theme
        theme_key = request.form.get('theme', 'theme1')
        
        # Get the creator name from the form
        creator_name = request.form.get('creator_name', '')
        
        # Get presentation rules from the form
        presentation_rules = request.form.get('presentation_rules', '')
        
        if file and allowed_file(file.filename):
            try:
                filename = secure_filename(file.filename)
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(pdf_path)
                print(f"Saved PDF to: {pdf_path}")
                
                # Extract text and images from PDF
                Text_extract.extract_text_and_images_from_pdf(pdf_path, app.config['EXTRACT_FOLDER'])
                print("Extracted text and images from PDF")
                
                # Read the extracted content
                content = txt_to_vba.read_input_files(app.config['EXTRACT_FOLDER'])
                
                # Add presentation rules to the content if provided
                if presentation_rules:
                    content += f"\n\nPRESENTATION RULES:\n{presentation_rules}"
                    print(f"Added presentation rules: {presentation_rules}")
                    print("IMPORTANT: Applying specific presentation formatting rules to the content.")
                
                # Generate outline using Gemini
                num_content_slides = 6  # You can adjust this number as needed
                gemini_output = txt_to_vba.generate_outline_with_gemini(content, num_content_slides)
                
                # Generate VBA code with creator name
                slides = txt_to_vba.parse_gemini_output(gemini_output)
                vba_code = txt_to_vba.generate_vba_code(slides, creator_name)
                
                # Save the VBA code
                with open('create_presentation.vba', 'w', encoding='utf-8') as f:
                    f.write(vba_code)
                
                # Set output path for PowerPoint
                ppt_output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'generated_presentation.pptx')
                
                # Convert VBA to PowerPoint with selected theme
                ppt_path = vba_to_ppt.create_powerpoint(
                    vba_to_ppt.parse_vba_file('create_presentation.vba'),
                    ppt_output_path,
                    theme_key
                )
                print(f"PowerPoint path: {ppt_path}")
                
                if ppt_path and os.path.exists(ppt_path):
                    # Send the file
                    response = send_file(ppt_path, as_attachment=True)
                    
                    # Clean up after sending the file
                    cleanup_folders()
                    
                    return response
                else:
                    raise FileNotFoundError(f"Generated PowerPoint file not found: {ppt_path}")
            except Exception as e:
                error_message = f"An error occurred: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
                print(error_message)  # Print to console for debugging
                # Clean up even if there's an error
                cleanup_folders()
                return render_template('index.html', message=error_message, themes=themes)
    
    return render_template('index.html', themes=themes)

@app.route('/generate-from-topic', methods=['POST'])
def generate_from_topic():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'
        }
        for key, theme in THEMES.items()
    ]
    
    try:
        # Get form data
        topic = request.form.get('topic', '')
        details = request.form.get('details', '')
        theme_key = request.form.get('theme', 'theme1')
        creator_name = request.form.get('creator_name', '')
        
        # Get slide count (default to 8 if not provided)
        try:
            slide_count = int(request.form.get('slide_count', 8))
            # Ensure slide count is within reasonable limits
            slide_count = max(4, min(slide_count, 12))
        except ValueError:
            slide_count = 8
        
        # Get presentation rules from the form
        presentation_rules = request.form.get('presentation_rules', '')
        
        if not topic:
            return render_template('topic_generator.html', message='Error: No topic provided', themes=themes)
        
        print(f"Generating presentation on topic: {topic}")
        print(f"Additional details: {details}")
        print(f"Using theme: {theme_key}")
        print(f"Creator: {creator_name}")
        print(f"Slide count: {slide_count}")
        if presentation_rules:
            print(f"Presentation rules specified:")
            for line in presentation_rules.strip().split('\n'):
                print(f"  - {line.strip()}")
        else:
            print("No presentation rules specified, using default formatting")
        
        # Ensure extract directory exists
        os.makedirs(app.config['EXTRACT_FOLDER'], exist_ok=True)
        os.makedirs(os.path.join(app.config['EXTRACT_FOLDER'], 'images'), exist_ok=True)
        
        # Generate presentation content based on the topic
        gemini_output = generate_topic_content(topic, details, slide_count, presentation_rules)
        
        # Parse the output and generate VBA code
        slides = txt_to_vba.parse_gemini_output(gemini_output)
        vba_code = txt_to_vba.generate_vba_code(slides, creator_name)
        
        # Extract slide titles for image search
        slide_titles = [slide['title'] for slide in slides if slide['title']]
        
        # Fetch relevant images for the topic
        fetch_images_for_topic(topic, slide_titles, num_images=min(8, slide_count + 2))
        
        # Save the VBA code
        with open('create_presentation.vba', 'w', encoding='utf-8') as f:
            f.write(vba_code)
        
        # Set output path for PowerPoint
        output_filename = f"topic_{topic.replace(' ', '_')[:30]}.pptx"
        ppt_output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        # Create output directory if it doesn't exist
        os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
        
        # Convert VBA to PowerPoint with selected theme
        ppt_path = vba_to_ppt.create_powerpoint(
            vba_to_ppt.parse_vba_file('create_presentation.vba'),
            ppt_output_path,
            theme_key
        )
        
        if ppt_path and os.path.exists(ppt_path):
            # Send the file
            response = send_file(ppt_path, as_attachment=True)
            return response
        else:
            raise FileNotFoundError(f"Generated PowerPoint file not found: {ppt_path}")
            
    except Exception as e:
        error_message = f"An error occurred: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_message)  # Print to console for debugging
        return render_template('topic_generator.html', message=error_message, themes=themes)

@app.route('/topic-generator')
def topic_generator():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'
        }
        for key, theme in THEMES.items()
    ]
    return render_template('topic_generator.html', themes=themes)

if __name__ == '__main__':
    # Create necessary directories
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(EXTRACT_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Clean up any leftover files from previous runs
    cleanup_folders()
    
    app.run(debug=True)
