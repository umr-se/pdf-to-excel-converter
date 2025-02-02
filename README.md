# PDF to Excel Converter

## Overview

**PDF to Excel Converter** is a backend Python script developed on Google Colab. It utilizes OCR (Optical Character Recognition) to extract text from PDFs and accurately convert it into an Excel file with precise column detection.

## Features

- Extracts text from PDFs using OCR
- Converts extracted text into structured Excel format
- Ensures precise column detection for better data organization
- Developed in Python on Google Colab for easy execution

## Prerequisites

Before running the script, ensure you have the following dependencies installed:

```bash
pip install pytesseract pdf2image pandas openpyxl
```

Additionally, install Tesseract OCR and ensure it is configured correctly:

- **Windows**: Download and install from [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux/macOS**: Install via package manager:
  ```bash
  sudo apt install tesseract-ocr   # For Ubuntu/Debian
  brew install tesseract           # For macOS
  ```

## Usage

1. Clone the repository:
   ```bash
   git clone https://github.com/umr-se/pdf-to-excel-converter
   cd pdf-to-excel-converter
   ```
2. Upload your PDF file to Google Colab.
3. Run the script in Google Colab.
4. The extracted text will be saved as an Excel file.

## Example Code

```python
!pip install pandas tabula-py openpyxl
!pip install pdfplumber
!pip install pytesseract pillow pandas openpyxl
!pip install paddlepaddle paddleocr
!sudo apt install tesseract-ocr

import re
from paddleocr import PaddleOCR
import pdfplumber
import numpy as np
from openpyxl import Workbook

# Initialize PaddleOCR
ocr = PaddleOCR(use_angle_cls=True, lang='en')

def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF using PaddleOCR.
    """
    extracted_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Convert the page to an image
            img = page.to_image()
            img_pil = img.original  # Get the PIL Image object

            # Convert the PIL Image to a NumPy array
            img_np = np.array(img_pil)

            # Pass the NumPy array to PaddleOCR for text extraction
            ocr_result = ocr.ocr(img_np, cls=True)
            for result in ocr_result[0]:
                extracted_text += result[1][0] + '\n'

    return extracted_text

def process_extracted_text_to_excel(text, output_file='output.xlsx'):
    """
    Process extracted text and save it into an Excel file with specified fields.
    """
    # Define the fields (column headers)
    fields = ["Lot no.Associated lots", "Unit no.", "U/E Lot no.Associated lots", "Unit no.", "U/E"]

    # Prepare data storage
    rows = []
    lines = text.split("\n")

    # Locate the start of data after "U/E"
    data_start_index = 0
    for i, line in enumerate(lines):
        if "U/E" in line:  # Look for the last "U/E" header
            data_start_index = i + 1

    # Process data starting after the last "U/E" header
    data_lines = lines[data_start_index:]
    cleaned_lines = [line.strip() for line in data_lines if line.strip()]  # Remove empty lines

    # Chunk data into groups of 5 to match the fields
    for i in range(0, len(cleaned_lines), 5):
        chunk = cleaned_lines[i:i + 5]
        if len(chunk) == 5:  # Ensure full rows
            rows.append(chunk)

    # Write data to Excel
    workbook = Workbook()
    sheet = workbook.active

    # Write headers
    sheet.append(fields)

    # Write rows
    for row in rows:
        sheet.append(row)

    # Save the Excel file
    workbook.save(output_file)
    print(f"Data has been saved to {output_file}")

def main(pdf_path):
    # Step 1: Extract text from the PDF
    extracted_text = extract_text_from_pdf(pdf_path)

    # Step 2: Process the text and save it to Excel
    process_extracted_text_to_excel(extracted_text)

# Main execution
pdf_path = '/content/test.pdf'  # Replace with your PDF file path
main(pdf_path)
```
## Contributing

Contributions are welcome! Feel free to submit issues or pull requests to improve the project.

## Contact

For any questions or feedback, reach out via GitHub Issues.


![img1](https://github.com/user-attachments/assets/eeb86772-78b6-42f7-b911-74dd65b4cc38)
![excel](https://github.com/user-attachments/assets/bcec89e7-86fa-4835-9a9d-b17d5621d58b)

