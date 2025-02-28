import os
import traceback
import shutil
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

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'  # Update path format
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
        
        if file and allowed_file(file.filename):
            try:
                filename = secure_filename(file.filename)
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(pdf_path)
                print(f"Saved PDF to: {pdf_path}")
                
                # Extract text and images from PDF
                Text_extract.extract_text_and_images_from_pdf(pdf_path, app.config['EXTRACT_FOLDER'])
                print("Extracted text and images from PDF")
                
                # Generate VBA code
                txt_to_vba.main()
                print("Generated VBA code")
                
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

if __name__ == '__main__':
    # Create necessary directories
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(EXTRACT_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Clean up any leftover files from previous runs
    cleanup_folders()
    
    app.run(debug=True)
