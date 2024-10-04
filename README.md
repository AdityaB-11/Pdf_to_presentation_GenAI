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
   git clone https://github.com/your-username/pdf-to-ppt.git
   cd pdf-to-ppt
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

1. Place your input PDF files in the `extract` folder.

2. Run the script:
   ```
   python txt_to_vba.py
   ```

3. The script will generate a VBA file named `create_presentation.vba`.

4. Open Microsoft PowerPoint and press Alt + F11 to open the VBA editor.

5. Import the generated `create_presentation.vba` file and run the macro to create your presentation.

## Configuration

You can customize the number of content slides by modifying the `num_content_slides` variable in the `main()` function of `txt_to_vba.py`.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Google Gemini AI for powering the content generation
<<<<<<< HEAD
- OpenAI for inspiration and guidance
=======
- OpenAI for inspiration and guidance
>>>>>>> c8013370e2d8cb28173c125da97c8fd24359e994
