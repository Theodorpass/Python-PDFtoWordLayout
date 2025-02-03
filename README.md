# Python-PDFtoWordLayout
# PDF to Word Conversion with Images

This Python script converts a PDF file into a Word document, extracting both the text and images. It uses the **PyMuPDF** library to extract the content and **python-docx** to generate the Word document. The text is placed in a table, and the images are inserted into the second column, preserving the layout of the original PDF.

## Requirements

To run this script, you need to install the following Python libraries:

```bash
pymupdf
python-docx
pillow
```

Alternatively, you can install all dependencies at once by using:

```bash
pip install -r requirements.txt
```

## Installation

1. Clone or download this repository to your local machine.

   ```bash
   git clone https://github.com/Theodorpass/PDF-to-Word-with-Images.git
   ```

2. Navigate to the project folder:

   ```bash
   cd PDF-to-Word-with-Images
   ```

3. Install the required dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Open the `pdf_to_word_with_images.py` script and modify the `pdf_file` and `word_file` paths to your specific files:

   ```python
   pdf_file = r'C:\path\to\your\input.pdf'  # Replace with your input PDF file path
   word_file = r'C:\path\to\your\output_with_images.docx'  # Replace with your desired output Word file path
   ```

2. Run the script:

   ```bash
   python pdf_to_word_with_images.py
   ```

3. After execution, check the destination folder for the newly created Word document containing the extracted text and images from the PDF.

## Error Logging

If any error occurs during the conversion, the script will create an **error_log.txt** file in the same directory as the output Word file. The log file will contain the error details and traceback information.

