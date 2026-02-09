#!/bin/bash

# PowerPoint Text Extractor Manager
# This script helps manage the folder structure and processing of PPTX files

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_SCRIPT="$SCRIPT_DIR/pptx_extractor.py"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

print_usage() {
    echo "Usage: $0 [COMMAND] [OPTIONS]"
    echo ""
    echo "Commands:"
    echo "  setup [DIR]            Create folder structure in specified directory (default: current)"
    echo "  process [DIR] [--keep] Process PPTX files (--keep leaves originals in input/)"
    echo "  clean [DIR]            Clean up processed and output folders"
    echo "  status [DIR]           Show status of folders and files"
    echo "  watch [DIR]            Watch input folder and auto-process new files"
    echo ""
    echo "Examples:"
    echo "  $0 setup ./my_presentations"
    echo "  $0 process"
    echo "  $0 process --keep"
    echo "  $0 status"
    echo "  $0 clean"
}

check_python_script() {
    if [[ ! -f "$PYTHON_SCRIPT" ]]; then
        echo -e "${RED}Error: Python script not found at $PYTHON_SCRIPT${NC}"
        echo "Make sure pptx_extractor.py is in the same directory as this script."
        exit 1
    fi
}

get_python_command() {
    # Try different Python commands to find what works
    if command -v python3 &> /dev/null; then
        echo "python3"
    elif command -v py &> /dev/null; then
        echo "py"
    elif command -v python &> /dev/null; then
        echo "python"
    else
        echo -e "${RED}Error: No Python installation found${NC}"
        echo "Please install Python or ensure it's in your PATH"
        exit 1
    fi
}

setup_folders() {
    local dir=${1:-.}
    local python_cmd=$(get_python_command)
    echo -e "${BLUE}Setting up folder structure in: $dir${NC}"
    
    $python_cmd "$PYTHON_SCRIPT" --dir "$dir" --setup
    
    if [[ $? -eq 0 ]]; then
        echo -e "${GREEN}✅ Folder structure created successfully${NC}"
        echo -e "${YELLOW}💡 You can now place PPTX files in the 'input' folder${NC}"
    else
        echo -e "${RED}❌ Failed to create folder structure${NC}"
        exit 1
    fi
}

process_files() {
    local dir=${1:-.}
    local keep=${2:-}
    local python_cmd=$(get_python_command)
    echo -e "${BLUE}Processing PPTX files in: $dir${NC}"

    if [[ "$keep" == "--keep" ]]; then
        $python_cmd "$PYTHON_SCRIPT" --dir "$dir" --keep
    else
        $python_cmd "$PYTHON_SCRIPT" --dir "$dir"
    fi

    if [[ $? -eq 0 ]]; then
        echo -e "${GREEN}✅ Processing completed${NC}"
    else
        echo -e "${RED}❌ Processing failed${NC}"
        exit 1
    fi
}

show_status() {
    local dir=${1:-.}
    local abs_dir=$(realpath "$dir")
    
    echo -e "${BLUE}Status for directory: $abs_dir${NC}"
    echo ""
    
    # Check if folders exist
    for folder in "input" "processed" "output"; do
        local folder_path="$abs_dir/$folder"
        if [[ -d "$folder_path" ]]; then
            local count=$(find "$folder_path" -maxdepth 1 -type f | wc -l)
            echo -e "${GREEN}📁 $folder: exists ($count files)${NC}"
        else
            echo -e "${YELLOW}📁 $folder: missing${NC}"
        fi
    done
    
    echo ""
    
    # Show PPTX files in input
    local input_dir="$abs_dir/input"
    if [[ -d "$input_dir" ]]; then
        local pptx_count=$(find "$input_dir" -maxdepth 1 -iname "*.pptx" | wc -l)
        if [[ $pptx_count -gt 0 ]]; then
            echo -e "${YELLOW}📄 PPTX files ready to process:${NC}"
            find "$input_dir" -maxdepth 1 -iname "*.pptx" -exec basename {} \;
        else
            echo -e "${BLUE}📄 No PPTX files in input folder${NC}"
        fi
    fi
}

clean_folders() {
    local dir=${1:-.}
    local abs_dir=$(realpath "$dir")
    
    echo -e "${YELLOW}⚠️  This will delete all files in 'processed' and 'output' folders${NC}"
    read -p "Are you sure? (y/N): " -n 1 -r
    echo
    
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        rm -rf "$abs_dir/processed"/* 2>/dev/null
        rm -rf "$abs_dir/output"/* 2>/dev/null
        echo -e "${GREEN}✅ Cleaned processed and output folders${NC}"
    else
        echo -e "${BLUE}Operation cancelled${NC}"
    fi
}

watch_folder() {
    local dir=${1:-.}
    local input_dir="$dir/input"

    echo -e "${BLUE}👀 Watching $input_dir for new PPTX files...${NC}"
    echo -e "${YELLOW}Press Ctrl+C to stop${NC}"

    # Create input folder if it doesn't exist
    mkdir -p "$input_dir"

    # macOS: use fswatch
    if [[ "$(uname)" == "Darwin" ]]; then
        if ! command -v fswatch &> /dev/null; then
            echo -e "${RED}Error: fswatch not found. Install it with: brew install fswatch${NC}"
            exit 1
        fi
        fswatch -0 --event Created --event MovedTo "$input_dir" | while read -d '' event; do
            echo -e "${YELLOW}📥 New file detected, processing...${NC}"
            sleep 2
            process_files "$dir"
            echo -e "${BLUE}👀 Continuing to watch...${NC}"
        done
    # Linux: use inotifywait
    elif command -v inotifywait &> /dev/null; then
        while inotifywait -e moved_to,create "$input_dir" 2>/dev/null; do
            echo -e "${YELLOW}📥 New file detected, processing...${NC}"
            sleep 2
            process_files "$dir"
            echo -e "${BLUE}👀 Continuing to watch...${NC}"
        done
    else
        echo -e "${RED}Error: No file watcher found.${NC}"
        echo "macOS: brew install fswatch"
        echo "Ubuntu/Debian: sudo apt-get install inotify-tools"
        echo "CentOS/RHEL: sudo yum install inotify-tools"
        exit 1
    fi
}

# Main script logic
case "$1" in
    setup)
        check_python_script
        setup_folders "$2"
        ;;
    process)
        check_python_script
        process_files "$2" "$3"
        ;;
    status)
        show_status "$2"
        ;;
    clean)
        clean_folders "$2"
        ;;
    watch)
        check_python_script
        watch_folder "$2"
        ;;
    "")
        print_usage
        ;;
    *)
        echo -e "${RED}Unknown command: $1${NC}"
        echo ""
        print_usage
        exit 1
        ;;
esac