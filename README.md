# PPTScanner: AI Presentation Consistency Checker

PPTScanner is a command-line Python tool that analyzes PowerPoint presentations (`.pptx` files) to find factual, numerical, and logical inconsistencies across slides. It uses a combination of text extraction, Optical Character Recognition (OCR) for images, and a powerful Large Language Model (Google's Gemini 1.5 Flash) to perform its analysis.

This tool is designed for consultants, analysts, and anyone who needs to ensure the accuracy and coherence of their presentation decks before a critical meeting.

## Key Features

* **Comprehensive Text Extraction**: Extracts text from all standard PowerPoint elements, including text boxes, shapes, and tables.
* **Image-to-Text (OCR)**: Automatically detects and extracts text from any images embedded within the slides, ensuring no data is missed.
* **AI-Powered Analysis**: Leverages the Gemini 1.5 Flash API to understand context and identify a wide range of inconsistencies:
    * **Numerical Conflicts**: Mismatched revenue figures, incorrect totals, conflicting percentages.
    * **Textual Contradictions**: Opposing claims or statements (e.g., "high growth market" vs. "stagnant industry").
    * **Timeline Mismatches**: Conflicting dates, project phases, or launch forecasts.
* **Clear & Structured Output**: Reports findings in a clean, easy-to-read format in the terminal, referencing specific slide numbers and explaining each issue.
* **Generalizable**: Designed to work on any `.pptx` presentation without prior configuration.

## How It Works

1.  **PPTX Parsing (`python-pptx`)**: The script iterates through each slide in the presentation. It systematically extracts text from all shapes, text frames, and table cells.
2.  **Image Extraction & OCR (`Pillow` & `pytesseract`)**: When an image is found, it's extracted in memory. `Pillow` processes the image, and `pytesseract` performs OCR to convert any text within the image into a string.
3.  **Data Consolidation**: All extracted text (from shapes and OCR) for each slide is aggregated into a single text block, clearly marked with its slide number (e.g., `--- Slide 3 ---`).
4.  **AI Analysis (`Gemini 1.5 Flash`)**: This consolidated text from all slides is sent to the Gemini API with a carefully crafted prompt. The prompt instructs the AI to act as an expert consultant and identify specific types of inconsistencies.
5.  **Formatted Output**: The AI's response is then parsed and printed to the terminal in a structured, human-readable report.

## Setup and Installation

Follow these steps to get PPTScanner running on your local machine.

### 1. Prerequisites

* **Python 3.8+**
* **Tesseract-OCR**: This is a mandatory dependency for reading text from images.
    * **Windows**: Download the installer from the [Tesseract at UB Mannheim repository](https://github.com/UB-Mannheim/tesseract/wiki). **Important:** During installation, make sure to add Tesseract to your system's `PATH` environment variable.
    * **macOS**: `brew install tesseract`
    * **Linux (Debian/Ubuntu)**: `sudo apt-get update && sudo apt-get install tesseract-ocr`

### 2. Clone the Repository

```bash
git clone [https://github.com/gurnooriitd/PPTScanner.git](https://github.com/gurnooriitd/PPTScanner.git)
cd PPTScanner
```

### 3. Set Up a Python Environment

It's highly recommended to use a virtual environment to avoid conflicts with other projects.

```bash
# Create the virtual environment
python -m venv venv

# Activate it
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

### 4. Install Dependencies

Install all the required Python packages from the `requirements.txt` file.

```bash
pip install -r requirements.txt
```

### 5. Configure Your API Key

The script requires a Google Gemini API key.

1.  Create a file named `.env` in the root of the project directory.
2.  Add your API key to this file in the following format:
    ```
    GEMINI_API_KEY="YOUR_API_KEY_HERE"
    ```
    You can get a free key from [Google AI Studio](https://aistudio.google.com/app/apikey).

## Usage

Run the script from your terminal, passing the full path to the PowerPoint file you want to analyze. **Remember to wrap the path in quotes if it contains spaces.**

```bash
python main.py "C:\path\to\your\presentation.pptx"
```

### Example Output

```
[+] Starting analysis of 'Sample-Deck.pptx'...
[+] Extracting text from 5 slides...
    - Extracted text from Slide 1
    - Extracted text and 1 image(s) from Slide 2
    - Extracted text from Slide 3
    - Extracted text and 1 image(s) from Slide 4
    - Extracted text from Slide 5
[+] Extraction complete. Consolidating text for analysis.
[+] Sending data to AI for inconsistency analysis...
[+] Analysis Complete.

========================================
      AI Inconsistency Report
========================================

Here are the inconsistencies found in the presentation:

1.  **Inconsistency: Conflicting Revenue Projections**
    * **Location:** Slide 2 vs. Slide 4.
    * **Conflict:** Slide 2 states, "2024 Revenue Projections: $15M," while the chart on Slide 4 shows the "Projected Rev. 2024" as "$12M."
    * **Explanation:** The revenue projection for 2024 is stated as two different values on two separate slides, which is a direct numerical contradiction.

2.  **Inconsistency: Contradictory Market Position Statement**
    * **Location:** Slide 3 vs. Slide 5.
    * **Conflict:** Slide 3 claims, "Our key advantage is operating in a market with few competitors," whereas Slide 5 states, "Navigating a highly competitive landscape is our main challenge."
    * **Explanation:** The presentation makes contradictory claims about the competitive environment.

## Limitations and Future Improvements

* **OCR Accuracy**: The quality of OCR is dependent on image resolution and text clarity. Stylized fonts, complex backgrounds, or low-quality images can lead to inaccurate text extraction.
* **No Visual Chart Interpretation**: The tool **does not visually interpret graphs or charts**. It can only read text labels, titles, or data points explicitly written as text on the chart image. It cannot, for example, understand the trend of a line graph unless that trend is also described in text.
* **Contextual Nuance**: While Gemini is powerful, it may miss highly nuanced, industry-specific inconsistencies or occasionally "hallucinate" a minor issue. The output should always be reviewed by a human.
* **Scalability**: For extremely large presentations (200+ slides), the amount of text sent to the API could become very large. While Gemini 1.5 Flash has a large context window, performance might slow down.
