📌 AI-Powered PPTX Inconsistency Checker
Overview
This tool analyzes multi-slide PowerPoint presentations (.pptx) and associated slide images (.jpeg, .jpg, .png) to detect factual or logical inconsistencies.
It uses:

python-pptx to extract text from .pptx files

pytesseract to OCR text from slide images

Google Gemini 2.5 Flash for AI-driven comparison and reasoning

✨ Features
PPTX + Image Support: Extracts facts from both text layers and image-only slides.

AI-Powered Detection: Uses Gemini 2.5 Flash to identify:

Conflicting numbers (e.g., $2M vs $3M)

Contradictory textual claims (e.g., “2x faster” vs “3x faster”)

Timeline/date mismatches

Auto Image Detection: Automatically finds all slide images in the same folder as the PPTX.

Structured Output:

Terminal report for human readability

inconsistencies.json for programmatic use

Scalable: Works on small or large decks

📦 Installation
1. Install Python packages
bash
Copy code
pip install python-pptx pillow pytesseract google-generativeai
2. Install Tesseract OCR
macOS (Homebrew):

bash
Copy code
brew install tesseract
Ubuntu/Debian:

bash
Copy code
sudo apt update
sudo apt install tesseract-ocr
Windows:
Download from: https://github.com/UB-Mannheim/tesseract/wiki

🔑 Gemini API Key Setup
Get a free API key from: https://aistudio.google.com/app/apikey

Set it as an environment variable:

macOS/Linux:

bash
Copy code
export GEMINI_API_KEY="your_api_key_here"
Windows (CMD):

cmd
Copy code
set GEMINI_API_KEY=your_api_key_here
Windows (PowerShell):

powershell
Copy code
$env:GEMINI_API_KEY="your_api_key_here"
🚀 Usage
Example:
bash
Copy code
python3 pptx_inconsistency_checker.py /path/to/NoogatAssignment.pptx
What happens:

Extracts text from the PPTX

Auto-detects and OCRs all .jpeg/.jpg/.png files in the same folder

Sends extracted facts to Gemini 2.5 Flash

Prints inconsistencies in Terminal

Saves full JSON to inconsistencies.json

📄 Sample Output
Terminal:

pgsql
Copy code
=== Inconsistencies Found ===
1. [NUMERIC] Slides: 3, Slide3.jpeg, 4, Slide4.jpeg, 5, Slide5.jpeg → Declared total 50 hours vs sum of components 40 hours.
2. [NUMERIC] Slides: 1, Slide1.jpeg, 2, Slide2.jpeg → '2x faster' vs '3x faster'.
3. [NUMERIC] Slides: 1, Slide1.jpeg, 2, Slide2.jpeg → $2M vs $3M savings.
4. [NUMERIC] Slides: 1, Slide1.jpeg, 2, Slide2.jpeg, 7, Slide7.jpeg → 15 mins vs 20 mins saved per slide.
inconsistencies.json:

json
Copy code
[
  {
    "slides": [1, "Slide1.jpeg", 2, "Slide2.jpeg"],
    "type": "numeric",
    "description": "Slide 1 states '2x faster', Slide 2 states '3x faster deck creation speed'."
  }
]
⚠ Limitations
Relies on Gemini API availability and quota limits.

OCR quality depends on Tesseract’s accuracy with the slide images.

Numeric/textual matching is handled by AI — false positives/negatives possible.

.ppt (old format) is not supported — must be .pptx.

🛠 Tech Stack
Python 3.8+

python-pptx — PPTX text extraction

Pillow + pytesseract — Image OCR

google-generativeai — Gemini 2.5 Flash API client

📬 Author
Banoth Lokesh
For internship assignment submission — AI-enabled inconsistency detection for presentations.
