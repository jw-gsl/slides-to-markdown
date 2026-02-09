#!/usr/bin/env python3

from pptx import Presentation
import os
import shutil
import argparse
from pathlib import Path

def extract_text_from_shape(shape):
    text_items = []

    # Standard text
    if hasattr(shape, "text") and shape.text.strip():
        text_items.append(shape.text.strip())

    # Tables
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    text_items.append(cell_text)

    # Grouped shapes
    if shape.shape_type == 6:  # GROUP shape type
        for sub_shape in shape.shapes:
            text_items.extend(extract_text_from_shape(sub_shape))

    return text_items

def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    text_runs = []

    for slide_number, slide in enumerate(prs.slides, start=1):
        text_runs.append(f"\n--- Slide {slide_number} ---\n")
        for shape in slide.shapes:
            text_runs.extend(extract_text_from_shape(shape))

    return "\n\n".join(text_runs)

def create_folder_structure(base_dir, quiet=False):
    """Create the necessary folder structure"""
    folders = {
        'input': Path(base_dir) / 'input',
        'processed': Path(base_dir) / 'processed', 
        'output': Path(base_dir) / 'output'
    }
    
    for folder_name, folder_path in folders.items():
        if not folder_path.exists():
            folder_path.mkdir(parents=True, exist_ok=True)
            if not quiet:
                print(f"📁 Created folder: {folder_path}")
        elif not quiet:
            print(f"📁 Verified folder: {folder_path}")
    
    return folders

def process_pptx_files(base_dir):
    """Process all PPTX files in the input directory"""
    folders = create_folder_structure(base_dir, quiet=True)  # Quiet mode for processing
    
    input_dir = folders['input']
    processed_dir = folders['processed']
    output_dir = folders['output']
    
    # Find all PPTX files in input directory (case-insensitive)
    pptx_files = list(input_dir.glob('*.[Pp][Pp][Tt][Xx]'))

    # Filter out temporary Office files (start with ~$)
    pptx_files = [f for f in pptx_files if not f.name.startswith('~$')]

    if not pptx_files:
        print(f"❌ No PPTX files found in {input_dir}")
        return

    print(f"🔍 Found {len(pptx_files)} PPTX file(s) to process")

    for pptx_file in pptx_files:
        try:
            print(f"\n📄 Processing: {pptx_file.name}")
            
            # Extract text
            extracted_text = extract_text_from_pptx(pptx_file)
            
            # Create output text file
            base_name = pptx_file.stem
            output_file = output_dir / f"{base_name}.txt"
            
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(extracted_text)
            
            print(f"✅ Text extracted to: {output_file}")
            
            # Move processed file
            processed_file = processed_dir / pptx_file.name
            shutil.move(str(pptx_file), str(processed_file))
            print(f"📦 Moved to processed: {processed_file}")
            
        except Exception as e:
            print(f"❌ Error processing {pptx_file.name}: {str(e)}")

def main():
    parser = argparse.ArgumentParser(description='Extract text from PowerPoint files')
    parser.add_argument('--dir', '-d', default='.', 
                       help='Base directory for folder structure (default: current directory)')
    parser.add_argument('--setup', action='store_true',
                       help='Only create folder structure without processing files')
    
    args = parser.parse_args()
    
    base_dir = Path(args.dir).resolve()
    print(f"🎯 Working in directory: {base_dir}")
    
    if args.setup:
        create_folder_structure(base_dir)
        print("\n📋 Folder structure created. Place your PPTX files in the 'input' folder and run without --setup flag.")
    else:
        process_pptx_files(base_dir)

if __name__ == "__main__":
    main()
