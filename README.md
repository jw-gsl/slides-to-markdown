# Slides to Markdown

A Python tool for extracting text content from PowerPoint (.pptx) files into clean, readable text files. Drop your slides into the `input/` folder, run the script, and get extracted text in `output/`.

## Features

- Extracts text from slides, tables, and grouped shapes
- Batch processes multiple PPTX files at once
- Automatically organizes files into input, processed, and output folders
- Generates well-formatted text with slide separators
- Continues processing even if individual files fail

## Setup

### macOS

1. **Clone the repo**

   ```bash
   git clone https://github.com/jw-gsl/slides-to-markdown.git
   cd slides-to-markdown
   ```

2. **Check Python is installed** (macOS ships with Python 3 on recent versions)

   ```bash
   python3 --version
   ```

   If Python is not installed, install it via [Homebrew](https://brew.sh):

   ```bash
   brew install python
   ```

3. **Install the dependency**

   ```bash
   pip3 install python-pptx
   ```

4. **Create the folder structure**

   ```bash
   python3 pptx_extractor.py --setup
   ```

5. **Run it** — drop `.pptx` files into `input/` then:

   ```bash
   python3 pptx_extractor.py
   ```

   You can also use the shell helper:

   ```bash
   chmod +x pptx_manager.sh
   ./pptx_manager.sh process
   ```

### Windows

1. **Clone the repo**

   ```powershell
   git clone https://github.com/jw-gsl/slides-to-markdown.git
   cd slides-to-markdown
   ```

2. **Install Python** (if not already installed)

   Download and install from [python.org](https://www.python.org/downloads/). During installation, make sure to check **"Add Python to PATH"**.

   Verify it works:

   ```powershell
   python --version
   ```

3. **Install the dependency**

   ```powershell
   pip install python-pptx
   ```

4. **Create the folder structure**

   ```powershell
   python pptx_extractor.py --setup
   ```

5. **Run it** — drop `.pptx` files into `input\` then:

   ```powershell
   python pptx_extractor.py
   ```

## Usage

### Basic workflow

1. Place your `.pptx` files in the `input/` folder
2. Run `python3 pptx_extractor.py` (or `python` on Windows)
3. Find extracted text in `output/`
4. Originals are automatically moved to `processed/`

### Command line options

| Option    | Short | Description                                      |
|-----------|-------|--------------------------------------------------|
| `--dir`   | `-d`  | Specify base directory (default: current directory) |
| `--setup` |       | Create folder structure only, don't process files |
| `--help`  | `-h`  | Show help message                                |

### Shell manager (macOS/Linux)

The `pptx_manager.sh` script provides additional commands:

```bash
./pptx_manager.sh setup      # Create folder structure
./pptx_manager.sh process    # Process PPTX files
./pptx_manager.sh status     # Show folder status and file counts
./pptx_manager.sh clean      # Remove processed and output files
```

## Project Structure

```
slides-to-markdown/
├── pptx_extractor.py    # Main extraction script
├── pptx_manager.sh      # Shell helper (macOS/Linux)
├── input/               # Place .pptx files here
├── processed/           # Originals moved here after extraction
└── output/              # Extracted text files
```

## Output Format

```
--- Slide 1 ---

Welcome to Our Presentation
Introduction and Overview

--- Slide 2 ---

Key Points:
- Point one
- Point two
```

## Supported Content

- Text boxes and shapes
- Slide titles and content
- Bullet points and numbered lists
- Tables (all cells)
- Grouped shapes

## Troubleshooting

**"No PPTX files found"** — Make sure files are in the `input/` folder, not the root directory.

**"Module not found"** — Run `pip install python-pptx` (or `pip3` on macOS).

**"Permission denied"** — Close PowerPoint if the files are open, and check folder permissions.

**Note:** Only `.pptx` files are supported (PowerPoint 2007+). Older `.ppt` files must be converted first.
