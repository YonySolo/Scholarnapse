# Scholarnapse

A Python tool that reads PowerPoint (.pptx) files and generates a narrative synopsis of the presentation content.

## What It Does

Scholarnapse extracts content from any PowerPoint file and produces a readable, flowing summary — like a book synopsis for your slides. It identifies key sections, groups related slides together, and writes a connected narrative instead of just listing bullet points.

### Features

- **Full Content Extraction** — pulls slide titles, body text, images, tables, and speaker notes from any `.pptx` file
- **Smart Slide Grouping** — merges consecutive slides with the same title into unified sections
- **Narrative Synopsis** — generates a flowing summary with transition phrases that reads naturally
- **Detailed Breakdown** — optional mode with slide-by-slide content and word frequency analysis
- **Automatic Filtering** — skips title slides, activity prompts, and filler content
- **Dual Output** — saves both a summary and detailed `.txt` file

## Installation

1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/Scholarnapse.git
cd Scholarnapse
```

2. Create a virtual environment and install dependencies:
```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # Mac/Linux
pip install -r requirements.txt
```

## Usage

Place a `.pptx` file in the project folder, then update the filepath in `main.py` and run:

```bash
python main.py
```

This will:
- Print the synopsis to the terminal
- Save `_synopsis.txt` (narrative summary)
- Save `_detailed.txt` (full slide breakdown + key topics + synopsis)

## Example Output

```
SYNOPSIS
==================================================

This presentation on "Week 2 – Research Skills" covers 25 slides
across the following topics.
The main topics covered are: Lesson Overview, Research Skills,
Evaluating sources, Starting your research.

It then covers lesson overview. The research process, Evaluating
sources, Journal articles and Web sources...
```

## Project Structure

```
Scholarnapse/
├── main.py              # Main script with all functions
├── requirements.txt     # Python dependencies
└── README.md
```

## How It Works

The tool has four main functions:

1. **`extract_slides()`** — opens the `.pptx` file and loops through each slide's shapes to extract titles, text, images, tables, and speaker notes into a structured list of dictionaries.

2. **`group_slides_by_topic()`** — groups consecutive slides that share the same title into sections, reducing repetition in the output.

3. **`write_synopsis()`** — takes the grouped data and generates a narrative summary with an introduction, transition phrases between sections, and a conclusion with image/notes statistics.

4. **`generate_synopsis()`** — produces a detailed slide-by-slide breakdown with content previews and a word frequency analysis showing the top 15 key topics.

## Built With

- [python-pptx](https://python-pptx.readthedocs.io/) — PowerPoint file parsing
- Python standard library (`collections.Counter`, `sys`, `os`, `io`, `contextlib`)

## License

This project is open source and available for personal and educational use.
