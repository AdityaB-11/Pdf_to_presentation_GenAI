# PDF to PowerPoint Converter

This project converts PDF content into a PowerPoint presentation outline using Google's Gemini AI model and generates VBA code to create the presentation in Microsoft PowerPoint.

## Features

- Extracts text content from PDF files
- Generates a presentation outline using Google's Gemini AI
- Creates VBA code to automate PowerPoint presentation creation
- Handles multiple input files
- Customizable number of content slides

## Prerequisites

- Python 3.7+
- Google Cloud account with Gemini API access
- Microsoft PowerPoint (for running the generated VBA code)

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/AdityaB-11/pdf-to-ppt.git
   cd Pdf_to_presentation_GenAI
   ```

2. Install required Python packages:
   ```
   pip install -r requirements.txt
   ```

3. Set up your Google API key:
   - Create a `.env` file in the project root
   - Add your Google API key to the `.env` file:
     ```
     GOOGLE_API_KEY=your_actual_api_key_here
     ```

## Usage

  
1. Run the script:
   ```
   python app.py
   ```
2. You'll be then directed to localhost webui at http://127.0.0.1:5000 

   Upload the pdf you want to convert into presentation,it'll be downloaded to your device. 

## Configuration

You can customize the number of content slides by modifying the `num_content_slides` variable in the `main()` function of `txt_to_vba.py`.
 

 