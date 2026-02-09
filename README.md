# Slides to Markdown

A Python tool for extracting text content from PowerPoint (.pptx) files into markdown (.md) files. Drop your slides into the `input/` folder, run the script, and get markdown files in `output/`.

## Features

- Extracts text from slides, tables, and grouped shapes
- Renders tables as proper markdown tables
- Formats bullet points with indentation levels
- Uses slide titles as markdown headings
- Includes speaker notes as blockquotes
- Batch processes multiple PPTX files at once
- Automatically organizes files into input, processed, and output folders

## Setup

### macOS

1. **Install Git** (if not already installed)

   Check if Git is available:

   ```bash
   git --version
   ```

   If you see `command not found`, install it using one of these options:

   - **Xcode Command Line Tools** (recommended — will prompt automatically when you run `git`):

     ```bash
     xcode-select --install
     ```

   - **Homebrew**:

     ```bash
     brew install git
     ```

2. **Clone the repo**

   ```bash
   git clone https://github.com/jw-gsl/slides-to-markdown.git
   cd slides-to-markdown
   ```

3. **Check Python is installed** (macOS ships with Python 3 on recent versions)

   ```bash
   python3 --version
   ```

   If Python is not installed, install it via [Homebrew](https://brew.sh):

   ```bash
   brew install python
   ```

4. **Install dependencies**

   ```bash
   pip3 install -r requirements.txt
   ```

   > On corporate/managed Macs you may need: `pip3 install --user -r requirements.txt`

5. **Choose where to put your working folders**

   Run the interactive setup to pick where `input/`, `output/`, and `processed/` are created:

   ```bash
   python3 pptx_extractor.py --setup
   ```

   This will offer options like Desktop, Documents, or a custom path. Your choice is saved so you don't need to set it again.

6. **Run it** — drop `.pptx` files into `input/` then:

   ```bash
   python3 pptx_extractor.py
   ```

### Windows

1. **Install Git** (if not already installed)

   Check if Git is available:

   ```powershell
   git --version
   ```

   If it's not recognized, download and install [Git for Windows](https://git-scm.com/download/win). Use the default settings during installation.

   > **Don't want to install Git?** You can skip this step and [download the ZIP](https://github.com/jw-gsl/slides-to-markdown/archive/refs/heads/main.zip) instead — extract it and open the folder in your terminal.

2. **Clone the repo**

   ```powershell
   git clone https://github.com/jw-gsl/slides-to-markdown.git
   cd slides-to-markdown
   ```

3. **Install Python** (if not already installed)

   Download from [python.org](https://www.python.org/downloads/). During installation, check **"Add Python to PATH"**.

   Verify it works:

   ```powershell
   python --version
   ```

4. **Install dependencies**

   ```powershell
   python -m pip install -r requirements.txt
   ```

5. **Choose where to put your working folders**

   Run the interactive setup to pick where `input\`, `output\`, and `processed\` are created:

   ```powershell
   python pptx_extractor.py --setup
   ```

   This will offer options like Desktop, Documents, or a custom path. Your choice is saved so you don't need to set it again.

6. **Run it** — drop `.pptx` files into `input\` then:

   ```powershell
   python pptx_extractor.py
   ```

## Usage

### Basic workflow

1. Place your `.pptx` files in the `input/` folder
2. Run `python3 pptx_extractor.py` (or `python` on Windows)
3. Find markdown files in `output/`
4. Originals are automatically moved to `processed/`

### Command line options

| Option    | Short | Description                                              |
|-----------|-------|----------------------------------------------------------|
| `--setup` |       | Interactive setup — choose where working folders live    |
| `--dir`   | `-d`  | Override working directory for this run                   |
| `--keep`  | `-k`  | Keep originals in input/ instead of moving to processed/ |
| `--help`  | `-h`  | Show help message                                        |

### Helper scripts

**macOS/Linux** — `pptx_manager.sh`:

```bash
chmod +x pptx_manager.sh
./pptx_manager.sh process          # Process PPTX files
./pptx_manager.sh process --keep   # Process but keep originals
./pptx_manager.sh status           # Show folder status and file counts
./pptx_manager.sh clean            # Remove processed and output files
./pptx_manager.sh watch            # Auto-process new files (requires fswatch)
```

**Windows** — `pptx_manager.ps1`:

```powershell
.\pptx_manager.ps1 process         # Process PPTX files
.\pptx_manager.ps1 process -Keep   # Process but keep originals
.\pptx_manager.ps1 status          # Show folder status and file counts
.\pptx_manager.ps1 clean           # Remove processed and output files
```

## Project Structure

```
slides-to-markdown/
├── pptx_extractor.py    # Main extraction script
├── pptx_manager.sh      # Shell helper (macOS/Linux)
├── pptx_manager.ps1     # PowerShell helper (Windows)
├── requirements.txt     # Python dependencies
├── input/               # Place .pptx files here
├── processed/           # Originals moved here after extraction
└── output/              # Generated markdown files
```

## Output Format

The script uses slide titles as headings, renders tables as markdown, and appends speaker notes as blockquotes:

```markdown
# My Presentation

## Welcome & Overview

Introduction to the topic
Key objectives for today

## Q3 Revenue

| Quarter | Revenue | Growth |
| --- | --- | --- |
| Q1 | $100K | 5% |
| Q2 | $120K | 8% |

**Notes:**
> Remember to mention the partnership deal
```

## Supported Content

- Text boxes and shapes
- Slide titles (used as `##` headings)
- Bullet points with nesting levels
- Tables (rendered as markdown tables)
- Grouped shapes
- Speaker notes (as blockquotes)

## Troubleshooting

**"No PPTX files found"** — Make sure files are in the `input/` folder, not the root directory.

**"Module not found"** — Run `pip install -r requirements.txt` (or `pip3` on macOS).

**"Permission denied"** — Close PowerPoint if the files are open, and check folder permissions.

**Note:** Only `.pptx` files are supported (PowerPoint 2007+). Older `.ppt` files must be converted first.
