# AI-Powered PowerPoint Inconsistency Detector

A Python-based tool that uses the **Gemini 2.5 Flash API** to automatically find factual and logical inconsistencies across slides in a PowerPoint presentation. Designed for **consultants, analysts, and reviewers**, it automates a tedious and error-prone review process.

## ✨ Features

- **Dual-Mode Analysis**: Extracts both **text and tables** from `.pptx` files and analyzes **slide images** using multimodal AI.
- **Inconsistency Detection**: Flags:
  - Conflicting numerical data
  - Contradictory textual claims
  - Timeline mismatches
- **Clear Output**: Produces a **structured terminal report** with direct references to slide numbers for each issue.
- **Scalable**: Handles **large presentations** by iterating through each slide programmatically.

## 🚀 Getting Started

### Prerequisites

- Python **3.8+**
- A **Gemini API Key** (Get a free one from Google AI Studio)

### Installation

1. **Clone the repository**

```bash
git clone https://github.com/GarvitKhandelwal31/ppt-inconsistency-detector.git
cd ppt-inconsistency-detector
```

2. **Create and activate a virtual environment**

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

### Configuration

Create a `.env` file in the project's root directory and add your Gemini API key:

```env
GEMINI_API_KEY=YOUR_KEY_HERE
```

### Usage

1. Place your **PowerPoint file** (`.pptx`) and all **slide images** (`.jpeg`) inside the project folder.
2. Run the tool:

```bash
python inconsistency_detector.py NoogatAssignment.pptx images/
```

3. The results will be **printed directly in the terminal**.

## 📌 Limitations

- **Subtle Inconsistencies**: The model may miss very nuanced logical inconsistencies that require deep domain expertise.
- **Chart Interpretation**: Data extraction accuracy depends on the clarity of the chart — complex or overly stylized charts may cause errors.

## 🔮 Future Improvements

- Better structured **JSON output**
- Support for **PDFs and other formats**
- Improved **chart interpretation** capabilities

## 📜 License

This project is licensed under the **MIT License** — you are free to use, modify, and distribute it.