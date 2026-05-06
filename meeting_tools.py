import os
import re
import shutil
from pathlib import Path

try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    print("Note: pywin32 not installed. Outlook extraction will not be available.")
    print("Install with: pip install pywin32\n")

# ============================================================================
# CONFIGURATION - Edit these values for your setup
# ============================================================================

# Outlook Extractor Settings
OUTLOOK_ACCOUNT = "jasfper@amazon.com"
OUTLOOK_FOLDER = "Meeting Summaries"
EMAIL_LIMIT = None  # Set to a number to limit emails, or None for all

# File Processing Settings
PROCESSING_FOLDER = "./processing_folder"
OUTPUT_FOLDER = "./output_folder"

# File Organizer Categories
CATEGORIES = [
    "auctane",
    "pattern",
    "reese",
    "senthil",
    "santi"
]

# ============================================================================
# 1. OUTLOOK EMAIL EXTRACTOR
# ============================================================================

class OutlookEmailExtractor:
    def __init__(self):
        if not OUTLOOK_AVAILABLE:
            raise ImportError("pywin32 is required for Outlook extraction")
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    def list_folders(self, parent_folder=None, indent=0):
        """List all available folders"""
        if parent_folder is None:
            print("\nAvailable Outlook folders:")
            for account in self.outlook.Folders:
                print(f"\n{account.Name}")
                self.list_folders(account, 1)
        else:
            try:
                for folder in parent_folder.Folders:
                    print("  " * indent + f"- {folder.Name}")
                    if folder.Folders.Count > 0:
                        self.list_folders(folder, indent + 1)
            except:
                pass
    
    def get_folder(self, folder_path):
        """Get Outlook folder by path"""
        folders = folder_path.split('/')
        folder = None
        
        for account in self.outlook.Folders:
            if account.Name.lower() == folders[0].lower():
                folder = account
                folders = folders[1:]
                break
        
        if folder is None:
            raise ValueError(f"Account '{folders[0]}' not found")
        
        for folder_name in folders:
            found = False
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    folder = subfolder
                    found = True
                    break
            if not found:
                raise ValueError(f"Folder '{folder_name}' not found in '{folder.Name}'")
        
        return folder
    
    def extract_emails(self, folder_path, output_dir='emails', limit=None):
        """Extract emails from specified folder"""
        folder = self.get_folder(folder_path)
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        
        os.makedirs(output_dir, exist_ok=True)
        
        count = 0
        for message in messages:
            if limit and count >= limit:
                break
            
            try:
                subject = message.Subject
                sender = message.SenderName
                received = message.ReceivedTime
                body = message.Body
                
                timestamp = received.strftime('%Y%m%d_%H%M%S')
                safe_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_'))[:50]
                filename = f"{timestamp}_{safe_subject}.md"
                
                filepath = os.path.join(output_dir, filename)
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(f"From: {sender}\n")
                    f.write(f"Subject: {subject}\n")
                    f.write(f"Date: {received}\n")
                    f.write(f"\n{body}")
                
                count += 1
                print(f"Extracted: {subject}")
                
            except Exception as e:
                print(f"Error processing email: {e}")
                continue
        
        print(f"\nTotal emails extracted: {count}")
        return count

# ============================================================================
# 2. CLEAN MEETING FILES
# ============================================================================

def clean_file(file_path):
    """Remove base64 images and Zoom footer text from a file"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    content = re.sub(r'<data:image/png;base64[^>]*>', '', content)
    content = re.sub(r'To view your meeting summaries.*', '', content, flags=re.DOTALL)
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content.strip())
    
    print(f"Cleaned: {file_path}")

def clean_meeting_files(folder_path):
    """Clean all txt files in the specified folder"""
    if not os.path.exists(folder_path):
        print(f"Creating folder '{folder_path}'...")
        os.makedirs(folder_path, exist_ok=True)
    
    txt_files = [f for f in os.listdir(folder_path) if f.endswith('.md')]
    
    if not txt_files:
        print(f"No .md files found in '{folder_path}'")
        return
    
    print(f"Found {len(txt_files)} file(s). Processing...")
    
    for filename in txt_files:
        file_path = os.path.join(folder_path, filename)
        clean_file(file_path)
    
    print(f"\nDone! Processed {len(txt_files)} file(s)")

# ============================================================================
# 3. FIX FILENAMES
# ============================================================================

def extract_meeting_name(file_path):
    """Extract the full meeting name from the Subject line in the file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.startswith("Subject: Meeting summary: "):
                    meeting_name = line.replace("Subject: Meeting summary: ", "").strip()
                    return meeting_name
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return None

def sanitize_filename(name):
    """Remove or replace characters that are invalid in filenames"""
    sanitized = name.replace(' - ', '__')
    sanitized = sanitized.replace(' ', '_')
    sanitized = sanitized.replace("'", '')
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', sanitized)
    sanitized = sanitized.strip('. ')
    sanitized = re.sub(r'_+', '_', sanitized)
    return sanitized

def format_external_internal(text):
    """Format External/Internal keywords (including common typos) to be title-cased and wrapped in square brackets"""
    internal_typos = r'(?:internal|intenral|interna|interal|internla|interanl|itnernal)'
    external_typos = r'(?:external|extenral|externa|exteral|externla|exteranl|etxernal)'
    text = re.sub(r'^[_\s]*' + external_typos, '[External]', text, flags=re.IGNORECASE)
    text = re.sub(r'^[_\s]*' + internal_typos, '[Internal]', text, flags=re.IGNORECASE)
    return text

def fix_filenames(folder_path):
    """Process all files in the folder and rename them based on their content"""
    if not os.path.exists(folder_path):
        print(f"Creating folder '{folder_path}'...")
        os.makedirs(folder_path, exist_ok=True)
    
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    
    if not files:
        print(f"No files found in '{folder_path}'.")
        return
    
    print(f"Processing {len(files)} file(s) in '{folder_path}'...")
    print("=" * 80)
    print()
    
    for filename in files:
        file_path = os.path.join(folder_path, filename)
        meeting_name = extract_meeting_name(file_path)
        
        if meeting_name:
            base_name, ext = os.path.splitext(filename)
            prefix_match = re.match(r'^(\d{8}_\d{6})_Meeting\s*summary', base_name, re.IGNORECASE)
            
            if prefix_match:
                date_time = prefix_match.group(1)
                prefix = f"{date_time}_MeetingSummary_"
                formatted_meeting_name = format_external_internal(meeting_name)
                new_filename = sanitize_filename(f"{prefix}_{formatted_meeting_name}") + ext
            else:
                new_filename = sanitize_filename(meeting_name) + ext
            
            new_file_path = os.path.join(folder_path, new_filename)
            
            if filename != new_filename:
                counter = 1
                while os.path.exists(new_file_path):
                    new_filename = f"{sanitize_filename(meeting_name)}_{counter}{ext}"
                    new_file_path = os.path.join(folder_path, new_filename)
                    counter += 1
                
                try:
                    os.rename(file_path, new_file_path)
                    print(f"✓ RENAMED:")
                    print(f"  Old: {filename}")
                    print(f"  New: {new_filename}")
                    print(f"  Full meeting name: {meeting_name}")
                    print()
                except Exception as e:
                    print(f"✗ FAILED to rename '{filename}': {e}")
                    print()
        else:
            print(f"✗ WARNING: Could not find 'Subject: Meeting summary:' line in '{filename}'")
            print()
    
    print("=" * 80)
    print("\nDone!")

# ============================================================================
# 4. FILE ORGANIZER
# ============================================================================

def organize_files(source_folder, categories, output_folder=None):
    """Organize txt files into subfolders based on filename keywords"""
    source_path = Path(source_folder)
    output_path = Path(output_folder) if output_folder else source_path
    
    if not source_path.exists():
        print(f"Creating folder '{source_folder}'...")
        source_path.mkdir(parents=True, exist_ok=True)
    
    if output_folder and not output_path.exists():
        print(f"Creating output folder '{output_folder}'...")
        output_path.mkdir(parents=True, exist_ok=True)
    
    txt_files = list(source_path.glob("*.md"))
    
    if not txt_files:
        print(f"No .md files found in '{source_folder}'")
        return
    
    print(f"Found {len(txt_files)} file(s) to organize\n")
    
    uncategorized = []
    multiple_matches = []
    
    for txt_file in txt_files:
        filename_lower = txt_file.stem.lower()
        matches = [category for category in categories if category.lower() in filename_lower]
        
        if len(matches) == 0:
            uncategorized.append(txt_file.name)
            destination = output_path / txt_file.name
            shutil.move(str(txt_file), str(destination))
        elif len(matches) > 1:
            multiple_matches.append((txt_file.name, matches))
            destination = output_path / txt_file.name
            shutil.move(str(txt_file), str(destination))
        else:
            category = matches[0]
            category_folder = output_path / category
            category_folder.mkdir(exist_ok=True)
            
            destination = category_folder / txt_file.name
            shutil.move(str(txt_file), str(destination))
            print(f"Moved '{txt_file.name}' -> '{category}/'")
    
    if uncategorized:
        print(f"\n{len(uncategorized)} file(s) didn't match any category (moved to output folder):")
        for filename in uncategorized:
            print(f"  - {filename}")
    
    if multiple_matches:
        print(f"\n{len(multiple_matches)} file(s) matched multiple categories (moved to output folder):")
        for filename, matches in multiple_matches:
            print(f"  - {filename} (matches: {', '.join(matches)})")
    
    print("\nOrganization complete!")

# ============================================================================
# MAIN MENU
# ============================================================================

def print_menu():
    print("\n" + "=" * 60)
    print("MEETING TOOLS - All-in-One Utility")
    print("=" * 60)
    print("\n1. Extract emails from Outlook" + ("" if OUTLOOK_AVAILABLE else " (pywin32 not installed)"))
    print("2. Clean meeting files (remove base64 images and footers)")
    print("3. Fix filenames (rename based on content)")
    print("4. Organize files into category folders")
    print("5. Run all steps (2-4) in sequence")
    print("6. List Outlook folders" + ("" if OUTLOOK_AVAILABLE else " (pywin32 not installed)"))
    print("0. Exit")
    print("\nCurrent settings:")
    print(f"  - Processing folder: {PROCESSING_FOLDER}")
    print(f"  - Output folder: {OUTPUT_FOLDER}")
    print(f"  - Categories: {', '.join(CATEGORIES)}")
    if OUTLOOK_AVAILABLE:
        print(f"  - Outlook: {OUTLOOK_ACCOUNT}/{OUTLOOK_FOLDER}")
    print()

def main():
    while True:
        print_menu()
        choice = input("Select an option (0-6): ").strip()
        
        if choice == '0':
            print("\nExiting. Goodbye!")
            break
        
        elif choice == '1':
            if not OUTLOOK_AVAILABLE:
                print("\nError: pywin32 is not installed. Install with: pip install pywin32")
                continue
            print("\n--- Extracting Emails from Outlook ---")
            try:
                extractor = OutlookEmailExtractor()
                folder_path = f"{OUTLOOK_ACCOUNT}/{OUTLOOK_FOLDER}"
                extractor.extract_emails(folder_path, PROCESSING_FOLDER, EMAIL_LIMIT)
            except Exception as e:
                print(f"\nError: {e}")
        
        elif choice == '2':
            print("\n--- Cleaning Meeting Files ---")
            clean_meeting_files(PROCESSING_FOLDER)
        
        elif choice == '3':
            print("\n--- Fixing Filenames ---")
            fix_filenames(PROCESSING_FOLDER)
        
        elif choice == '4':
            print("\n--- Organizing Files ---")
            organize_files(PROCESSING_FOLDER, CATEGORIES, OUTPUT_FOLDER)
        
        elif choice == '5':
            print("\n--- Running All Steps ---")
            print("\nStep 1/3: Cleaning meeting files...")
            clean_meeting_files(PROCESSING_FOLDER)
            print("\nStep 2/3: Fixing filenames...")
            fix_filenames(PROCESSING_FOLDER)
            print("\nStep 3/3: Organizing files...")
            organize_files(PROCESSING_FOLDER, CATEGORIES, OUTPUT_FOLDER)
            print("\n✓ All steps completed!")
        
        elif choice == '6':
            if not OUTLOOK_AVAILABLE:
                print("\nError: pywin32 is not installed. Install with: pip install pywin32")
                continue
            print("\n--- Listing Outlook Folders ---")
            try:
                extractor = OutlookEmailExtractor()
                extractor.list_folders()
            except Exception as e:
                print(f"\nError: {e}")
        
        else:
            print("\nInvalid option. Please try again.")
        
        input("\nPress Enter to continue...")

if __name__ == "__main__":
    main()
