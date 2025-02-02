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
import pytesseract
from pdf2image import convert_from_path
import pandas as pd

# Convert PDF pages to images
images = convert_from_path("sample.pdf")

# Perform OCR on each page
text_data = [pytesseract.image_to_string(img) for img in images]

# Process text and convert to structured format
# (Add column detection logic here)

# Save to Excel
df = pd.DataFrame({"Extracted Text": text_data})
df.to_excel("output.xlsx", index=False)
```
## Contributing

Contributions are welcome! Feel free to submit issues or pull requests to improve the project.

## Contact

For any questions or feedback, reach out via GitHub Issues.


![img1](https://github.com/user-attachments/assets/eeb86772-78b6-42f7-b911-74dd65b4cc38)
![excel](https://github.com/user-attachments/assets/bcec89e7-86fa-4835-9a9d-b17d5621d58b)

