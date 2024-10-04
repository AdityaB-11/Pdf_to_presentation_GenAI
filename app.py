import os
import traceback
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import Text_extract
import txt_to_vba
import vba_to_ppt

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
EXTRACT_FOLDER = 'extract'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXTRACT_FOLDER'] = EXTRACT_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    print("Route accessed")  # Debug print
    if request.method == 'POST':
        print("POST request received")  # Debug print
        if 'file' not in request.files:
            return render_template('index.html', message='No file part')
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', message='No selected file')
        if file and allowed_file(file.filename):
            try:
                filename = secure_filename(file.filename)
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(pdf_path)
                print(f"File saved: {pdf_path}")  # Debug print
                
                # Extract text from PDF
                Text_extract.extract_text_from_pdf(pdf_path, app.config['EXTRACT_FOLDER'])
                print("Text extracted")  # Debug print
                
                # Generate VBA code
                txt_to_vba.main()
                print("VBA generated")  # Debug print
                
                # Convert VBA to PowerPoint
                ppt_path = vba_to_ppt.main()
                print(f"PowerPoint generated: {ppt_path}")  # Debug print
                
                if os.path.exists(ppt_path):
                    return send_file(ppt_path, as_attachment=True)
                else:
                    raise FileNotFoundError(f"Generated PowerPoint file not found: {ppt_path}")
            except Exception as e:
                error_message = f"An error occurred: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
                print(error_message)  # Print to console for debugging
                return render_template('index.html', message=error_message)
    return render_template('index.html')

if __name__ == '__main__':
    print("Creating directories")  # Debug print
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(EXTRACT_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    print("Starting Flask app")  # Debug print
    app.run(debug=True)
