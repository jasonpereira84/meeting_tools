# Meeting Tools - All-in-One Utility

A unified Python script that combines four meeting file management utilities into one convenient tool.

## Features

1. **Extract Emails from Outlook** - Pull meeting summaries from Outlook folders
2. **Clean Meeting Files** - Remove base64 images and Zoom footer text
3. **Fix Filenames** - Rename files based on their content (Subject line)
4. **Organize Files** - Sort files into category folders based on keywords
5. **Run All Steps** - Execute cleaning, fixing, and organizing in sequence

## Requirements

- Python 3.6 or higher
- `pywin32` (only needed for Outlook extraction)

## Installation

1. Install Python dependencies:
```bash
pip install pywin32
```

Note: If you don't need Outlook extraction, you can skip installing pywin32. The other features will still work.

## Configuration

Edit the configuration section at the top of `meeting_tools.py`:

```python
# Outlook Extractor Settings
OUTLOOK_ACCOUNT = "your_email@example.com"
OUTLOOK_FOLDER = "Meeting Summaries"
EMAIL_LIMIT = None  # Set to a number to limit emails, or None for all

# File Processing Settings
PROCESSING_FOLDER = "../processing_folder"

# File Organizer Categories
CATEGORIES = [
    "auctane",
    "pattern",
    "reese",
    "senthil"
]
```

## Usage

Run the script:
```bash
python meeting_tools.py
```

You'll see an interactive menu:

```
============================================================
MEETING TOOLS - All-in-One Utility
============================================================

1. Extract emails from Outlook
2. Clean meeting files (remove base64 images and footers)
3. Fix filenames (rename based on content)
4. Organize files into category folders
5. Run all steps (2-4) in sequence
6. List Outlook folders
0. Exit
```

## What Each Tool Does

### 1. Extract Emails from Outlook
- Connects to your Outlook application
- Extracts emails from a specified folder
- Saves each email as a text file with timestamp and subject
- Files are named: `YYYYMMDD_HHMMSS_Subject.txt`

### 2. Clean Meeting Files
- Removes `<data:image/png;base64...>` tags from text files
- Removes Zoom footer text starting from "To view your meeting summaries..."
- Processes all `.txt` files in the specified folder
- Overwrites files with cleaned content

### 3. Fix Filenames
- Reads the `Subject: Meeting summary:` line from each file
- Renames files to use the full meeting name from the content
- Preserves date/time prefixes (e.g., `20250804_140107_MeetingSummary_`)
- Formats External/Internal keywords as `[EXTERNAL]` or `[INTERNAL]`
- Sanitizes filenames (removes invalid characters, replaces spaces with underscores)
- Handles duplicate filenames by appending a counter

### 4. Organize Files
- Sorts txt files into subfolders based on keywords in filenames
- Creates category folders automatically
- Files matching exactly one keyword are moved to that category folder
- Files with no matches or multiple matches stay in the original location
- Case-insensitive matching

### 5. Run All Steps
Executes steps 2-4 in sequence:
1. Clean meeting files
2. Fix filenames
3. Organize files into categories

This is useful for processing a batch of newly extracted emails.

## Typical Workflow

1. **First time setup**: Configure the settings at the top of the script
2. **Extract emails**: Run option 1 to pull emails from Outlook
3. **Process files**: Run option 5 to clean, rename, and organize all files
4. **Repeat**: Run option 1 again for new emails, then option 5 to process them

## Tips

- Use option 6 to list all your Outlook folders if you're not sure of the exact folder name
- The script modifies files in place, so consider backing up important files first
- You can customize the `CATEGORIES` list to match your own keywords
- Set `EMAIL_LIMIT` to a small number (e.g., 10) when testing

## Troubleshooting

**"pywin32 not installed"**
- Install with: `pip install pywin32`
- Or skip Outlook extraction and use options 2-5 only

**"Account not found" or "Folder not found"**
- Use option 6 to list all available Outlook folders
- Update `OUTLOOK_ACCOUNT` and `OUTLOOK_FOLDER` with the exact names shown

**"Folder does not exist"**
- The script will automatically create the `PROCESSING_FOLDER` if it doesn't exist
- If you want to use a different location, update the path in the configuration

## Original Scripts

This tool combines functionality from:
- `outlook_email_extractor.py` - Extract emails from Outlook
- `clean_meeting_files.py` - Clean base64 images and footers
- `fix_filenames.py` - Rename files based on content
- `file_organizer.py` - Organize files into category folders
