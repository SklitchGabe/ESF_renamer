#!/usr/bin/env python3
import os
import time
import concurrent.futures
from pathlib import Path
import platform
import subprocess
from tqdm import tqdm
import sys
import shutil
import logging
import psutil
import argparse
import re
import PyPDF2
from langdetect import detect, LangDetectException

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pdf_conversion.log"),
        logging.StreamHandler()
    ]
)

def convert_with_word(input_file, output_file=None, retries=2):
    """Convert doc/docx to PDF using Microsoft Word (Windows only)"""
    if output_file is None:
        output_file = str(Path(input_file).with_suffix('.pdf'))
    
    # Only import win32com if we're using this function
    import win32com.client
    import pythoncom
    import time
    
    # Initialize COM in this thread
    pythoncom.CoInitialize()
    
    for attempt in range(retries + 1):
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # Don't show alerts
            
            # Set these additional properties for corporate environments
            word.Options.CheckGrammarAsYouType = False
            word.Options.CheckSpellingAsYouType = False
            
            # For OneDrive files, use a more robust approach
            if "OneDrive" in input_file:
                # Try different opening methods in case of issues
                try:
                    # Method 1: Open with ReadOnly flag to avoid lock issues
                    doc = word.Documents.Open(
                        os.path.abspath(input_file), 
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False
                    )
                    
                    # Try export method instead of SaveAs for OneDrive files
                    doc.ExportAsFixedFormat(
                        OutputFileName=os.path.abspath(output_file),
                        ExportFormat=17,  # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,    # wdExportOptimizeForPrint
                        CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
                        DocStructureTags=True
                    )
                    doc.Close(SaveChanges=False)
                    
                except Exception as e:
                    # If the first method fails, try a different approach
                    print(f"  First OneDrive method failed: {str(e)}")
                    print("  Trying alternative method...")
                    
                    # Force close any open documents
                    try:
                        for doc in word.Documents:
                            doc.Close(SaveChanges=False)
                    except:
                        pass
                    
                    # Method 2: Copy the file to temp directory first
                    import tempfile
                    import shutil
                    
                    temp_dir = tempfile.gettempdir()
                    temp_file = os.path.join(temp_dir, f"temp_{os.path.basename(input_file)}")
                    
                    try:
                        # Copy to temp location
                        shutil.copy2(input_file, temp_file)
                        
                        # Try with the temp file
                        doc = word.Documents.Open(temp_file)
                        doc.SaveAs(os.path.abspath(output_file), FileFormat=17)
                        doc.Close()
                        
                        # Clean up temp file
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    except Exception as temp_error:
                        raise Exception(f"Both OneDrive methods failed: {str(temp_error)}")
            else:
                # Standard approach for non-OneDrive files
                doc = word.Documents.Open(os.path.abspath(input_file))
                doc.SaveAs(os.path.abspath(output_file), FileFormat=17)  # 17 is PDF format
                doc.Close(SaveChanges=False)
                
            return output_file
            
        except Exception as e:
            if attempt < retries:
                print(f"  Attempt {attempt+1} failed for {os.path.basename(input_file)}: {str(e)}")
                # Wait before retrying
                time.sleep(3)  # Increased wait time for corporate environments
                
                # Force close any hanging Word instances before retrying
                try:
                    # First try to close Word gracefully if we still have a reference
                    if word:
                        try:
                            word.Quit()
                        except:
                            pass
                    
                    # Then use taskkill as a last resort
                    import subprocess
                    subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], 
                                  stdout=subprocess.DEVNULL, 
                                  stderr=subprocess.DEVNULL)
                    time.sleep(2)  # Give system more time to close Word
                except:
                    pass
            else:
                # All retries exhausted
                raise Exception(f"MS Word conversion failed after {retries+1} attempts: {str(e)}")
        finally:
            # Clean up COM resources
            if word:
                try:
                    word.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
    
    # This should not be reached, but just in case
    raise Exception("Unknown error in Word conversion")

def extract_project_id(pdf_path, max_pages=10):
    """
    Extract the first occurrence of a World Bank project ID from a PDF file.
    Project IDs are in the format P followed by 6 digits (e.g., P123456).
    Also handles cases where 0 is transcribed as O.
    
    Args:
        pdf_path: Path to the PDF file
        max_pages: Maximum number of pages to search (default: 10)
        
    Returns:
        The project ID if found, None otherwise
    """
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            # Limit the number of pages to search
            pages_to_search = min(len(reader.pages), max_pages)
            
            # Regular expression pattern for project ID (P followed by any 6 chars that could be digits or letter O)
            pattern = r'P[0-9O]{6}'
            
            # Search through pages
            for page_num in range(pages_to_search):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                # Find all matches in the page
                matches = re.findall(pattern, text)
                if matches:
                    # Get the first match and fix any O's that should be 0's
                    pid = matches[0]
                    # Replace letter 'O' with digit '0' in the project ID (starting after the P)
                    corrected_pid = 'P' + pid[1:].replace('O', '0')
                    logging.info(f"Found project ID: {pid}, corrected to: {corrected_pid}")
                    return corrected_pid
                    
        return None
        
    except Exception as e:
        logging.error(f"Error processing {pdf_path}: {str(e)}")
        return None

def get_unique_filename(base_path):
    """Generate a unique filename by adding a numeric suffix"""
    if not os.path.exists(base_path):
        return base_path
    
    directory = os.path.dirname(base_path)
    filename = os.path.basename(base_path)
    name, ext = os.path.splitext(filename)
    
    counter = 1
    while True:
        new_filename = f"{name}_{counter:02d}{ext}"
        new_path = os.path.join(directory, new_filename)
        if not os.path.exists(new_path):
            return new_path
        counter += 1

def process_file(file_path, output_dir, input_dir, rename_with_pid=True):
    """Process a single file conversion with error handling and optional PID renaming"""
    try:
        input_path = os.path.abspath(file_path)
        
        # Get the relative path more carefully
        rel_path = os.path.dirname(os.path.relpath(input_path, input_dir))
        
        if rel_path and rel_path != '.':
            target_dir = os.path.join(output_dir, rel_path)
            os.makedirs(target_dir, exist_ok=True)
        else:
            target_dir = output_dir
            
        output_name = Path(file_path).stem + ".pdf"
        output_path = os.path.join(target_dir, output_name)
        
        # Always check if output file already exists, regardless of the source
        if os.path.exists(output_path):
            logging.warning(f"File already exists, creating unique name: {output_path}")
            output_path = get_unique_filename(output_path)
            logging.info(f"Using unique name: {output_path}")
        
        # Log the paths to help debug
        logging.info(f"Converting: {input_path} -> {output_path}")
        
        # Ensure we're on Windows since Word is required
        if platform.system() != "Windows":
            raise Exception("Microsoft Word conversion requires Windows")
            
        # Convert using Word
        pdf_path = convert_with_word(input_path, output_path, retries=2)
        
        # If renaming with project ID is requested
        if rename_with_pid and pdf_path:
            # Extract project ID from the converted PDF
            project_id = extract_project_id(pdf_path)
            
            # If no project ID found in PDF content, check the filename
            if not project_id:
                logging.info(f"No project ID found in PDF content, checking filename: {os.path.basename(file_path)}")
                project_id = extract_project_id_from_filename(os.path.basename(file_path))
                if project_id:
                    logging.info(f"Found project ID in filename: {project_id}")
            
            if project_id:
                # Detect language (only if we successfully found a project ID)
                language_suffix = detect_language(pdf_path)
                
                # Create new filename with project ID and language suffix
                pid_filename = f"{project_id}_{language_suffix}.pdf"
                pid_path = os.path.join(target_dir, pid_filename)
                
                # Handle duplicate filenames
                if os.path.exists(pid_path):
                    counter = 1
                    while True:
                        temp_name = f"{project_id}_{language_suffix}_{counter:02d}.pdf"
                        temp_path = os.path.join(target_dir, temp_name)
                        if not os.path.exists(temp_path):
                            pid_path = temp_path
                            pid_filename = temp_name
                            break
                        counter += 1
                
                # Rename the file
                try:
                    os.rename(pdf_path, pid_path)
                    logging.info(f"Renamed to: {pid_filename}")
                    return (file_path, True, None, project_id)
                except Exception as e:
                    logging.error(f"Error renaming: {str(e)}")
                    return (file_path, True, None, None)
            else:
                logging.warning(f"No project ID found in PDF or filename: {pdf_path}")
                return (file_path, True, None, None)
        
        return (file_path, True, None, None)
    except Exception as e:
        return (file_path, False, str(e), None)

def copy_existing_pdfs(input_dir, output_dir, overwrite=False, rename_with_pid=True):
    """Copy all existing PDF files from input directory to output directory"""
    pdf_files = []
    for root, _, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return 0, {}
    
    print(f"\nCopying {len(pdf_files)} existing PDF files to output directory")
    
    copied = 0
    skipped = 0
    pid_mapping = {}  # To store file -> project ID mapping
    
    with tqdm(total=len(pdf_files), unit="file") as pbar:
        for pdf_file in pdf_files:
            # Create the relative path for maintaining folder structure
            rel_path = os.path.relpath(os.path.dirname(pdf_file), start=input_dir)
            if rel_path != '.':
                target_dir = os.path.join(output_dir, rel_path)
                os.makedirs(target_dir, exist_ok=True)
            else:
                target_dir = output_dir
            
            # Get the destination path
            dest_file = os.path.join(target_dir, os.path.basename(pdf_file))
            
            try:
                # If renaming with project ID is requested
                if rename_with_pid:
                    # Extract project ID from the PDF
                    project_id = extract_project_id(pdf_file)
                    
                    # If no project ID found in PDF content, check the filename
                    if not project_id:
                        logging.info(f"No project ID found in PDF content, checking filename: {os.path.basename(pdf_file)}")
                        project_id = extract_project_id_from_filename(os.path.basename(pdf_file))
                        if project_id:
                            logging.info(f"Found project ID in filename: {project_id}")
                    
                    if project_id:
                        # Detect language (only if we successfully found a project ID)
                        language_suffix = detect_language(pdf_file)
                        
                        # Create new filename with project ID and language suffix
                        pid_filename = f"{project_id}_{language_suffix}.pdf"
                        dest_file = os.path.join(target_dir, pid_filename)
                        
                        # Handle duplicate filenames
                        if os.path.exists(dest_file):
                            counter = 1
                            while True:
                                temp_name = f"{project_id}_{language_suffix}_{counter:02d}.pdf"
                                temp_path = os.path.join(target_dir, temp_name)
                                if not os.path.exists(temp_path):
                                    dest_file = temp_path
                                    pid_filename = temp_name
                                    break
                                counter += 1
                        
                        logging.info(f"Copying with project ID: {pdf_file} -> {dest_file}")
                        shutil.copy2(pdf_file, dest_file)
                        copied += 1
                        pid_mapping[dest_file] = project_id
                    else:
                        # No project ID found, use original filename
                        logging.warning(f"No project ID found in PDF or filename: {pdf_file}")
                        # Check if file already exists
                        if os.path.exists(dest_file):
                            unique_dest = get_unique_filename(dest_file)
                            shutil.copy2(pdf_file, unique_dest)
                            logging.info(f"Created unique filename: {unique_dest}")
                            copied += 1
                        else:
                            # No conflict, copy normally
                            shutil.copy2(pdf_file, dest_file)
                            copied += 1
                else:
                    # Standard copy without PID renaming
                    # Check if file already exists
                    if os.path.exists(dest_file):
                        # ALWAYS create a unique filename with numeric suffix
                        # regardless of the overwrite parameter
                        unique_dest = get_unique_filename(dest_file)
                        shutil.copy2(pdf_file, unique_dest)
                        logging.info(f"Created unique filename: {unique_dest}")
                        copied += 1
                    else:
                        # No conflict, copy normally
                        shutil.copy2(pdf_file, dest_file)
                        copied += 1
            except Exception as e:
                error_msg = f"Error copying {pdf_file}: {str(e)}"
                print(error_msg)
                logging.error(error_msg)
                skipped += 1
            
            pbar.update(1)
    
    print(f"PDF copying complete. Copied: {copied}, Skipped: {skipped}")
    return copied, pid_mapping

def convert_folder_to_pdf(rename_with_pid=True):
    """Convert all Word documents in a folder to PDF and optionally rename with Project IDs"""
    # Check if we're on Windows
    if platform.system() != "Windows":
        print("Error: This script requires Windows with Microsoft Word installed")
        return 1
    
    # Prompt user for the input folder path
    print("Please enter the path to the folder containing Word documents:")
    input_dir = input().strip()
    
    # Strip quotes if the user included them
    input_dir = input_dir.strip('"\'')
    
    # Validate input directory
    if not os.path.isdir(input_dir):
        print(f"Error: '{input_dir}' is not a valid directory")
        return 1
    
    # Prompt user for the output folder path
    print("Please enter the path to the output folder for PDF files:")
    output_dir = input().strip()
    
    # Strip quotes if the user included them
    output_dir = output_dir.strip('"\'')
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Find all .docx and .doc files
    word_files = []
    # Find all .pdf files
    pdf_files = []
    for root, _, files in os.walk(input_dir):
        for file in files:
            lower_file = file.lower()
            if lower_file.endswith('.docx') or lower_file.endswith('.doc'):
                word_files.append(os.path.join(root, file))
            elif lower_file.endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    # Dictionary to store project ID mappings
    project_id_mappings = {}
    
    if not word_files:
        print(f"No Word documents (.doc or .docx) found in {input_dir}")
        # Even if no Word files are found, we'll still copy PDFs
    else:
        # Get maximum number of workers for optimal performance
        # Use all available CPU cores
        max_workers = os.cpu_count()
        
        print(f"Found {len(word_files)} Word documents to convert")
        print(f"Using Microsoft Word for conversion with {max_workers} worker processes")
        
        # Initialize counters and timing
        start_time = time.time()
        successful = 0
        failed = 0
        
        # Determine batch size based on system memory
        batch_size = get_optimal_batch_size()
        
        print(f"Using batch size: {batch_size}")
        
        # Calculate total files to process
        total_files = len(word_files) + len(pdf_files)
        print(f"Total files to process: {total_files}")
        
        # Process in smaller batches to prevent memory issues
        for i in range(0, len(word_files), batch_size):
            batch = word_files[i:i+batch_size]
            
            print(f"\nProcessing batch {i//batch_size + 1} of {(len(word_files) + batch_size - 1) // batch_size} ({len(batch)} files)")
            
            # Clean up any existing Word processes before each batch
            try:
                subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], 
                              stdout=subprocess.DEVNULL, 
                              stderr=subprocess.DEVNULL)
                time.sleep(1)  # Give system time to close Word
            except:
                pass
                    
            # Process the current batch
            with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                # Submit jobs for this batch with input_dir as an additional argument
                future_to_file = {
                    executor.submit(process_file, file, output_dir, input_dir, rename_with_pid): file
                    for file in batch
                }
                
                # Process results
                with tqdm(total=len(batch), unit="file") as pbar:
                    for future in concurrent.futures.as_completed(future_to_file):
                        file_path, success, error, project_id = future.result()
                        if success:
                            successful += 1
                            if project_id:
                                project_id_mappings[file_path] = project_id
                        else:
                            failed += 1
                            print(f"Error converting {file_path}: {error}")
                        pbar.update(1)
        
        # Report results
        elapsed_time = time.time() - start_time
        files_per_second = len(word_files) / elapsed_time if elapsed_time > 0 else 0
        
        print(f"\nConversion complete in {elapsed_time:.2f} seconds ({files_per_second:.2f} files/sec)")
        print(f"Successfully converted: {successful}")
        print(f"Failed conversions: {failed}")
        
        # Add success rate report
        if successful + failed > 0:
            print(f"Success rate: {successful/(successful+failed)*100:.1f}%")
        else:
            print("Success rate: N/A (no files processed)")
    
    # Always set overwrite=False to ensure no files are overwritten
    copied_pdfs, pdf_pid_mappings = copy_existing_pdfs(input_dir, output_dir, overwrite=False, rename_with_pid=rename_with_pid)
    
    # Merge the project ID mappings
    project_id_mappings.update(pdf_pid_mappings)
    
    print(f"\nAll operations complete. Output files saved to: {output_dir}")
    
    # Verify the output directory contents
    output_files = []
    for root, _, files in os.walk(output_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                output_files.append(os.path.join(root, file))
    
    print(f"Actual PDF files in output directory: {len(output_files)}")
    
    if len(output_files) < (successful + copied_pdfs):
        print("WARNING: Some files may have been overwritten due to naming conflicts.")
    
    # Create a summary of project IDs found
    if rename_with_pid:
        project_ids = list(set(project_id_mappings.values()))
        print(f"\nFound {len(project_ids)} unique project IDs")
        if project_ids:
            print("Sample of project IDs found:")
            for pid in project_ids[:5]:  # Show first 5
                print(f"  - {pid}")
            if len(project_ids) > 5:
                print(f"  ...and {len(project_ids) - 5} more")
    
    return 0

def verify_pdf(pdf_path):
    """Verify that the created PDF is valid"""
    try:
        # Use PyPDF2 to check PDF validity
        import PyPDF2
        with open(pdf_path, 'rb') as file:
            try:
                pdf = PyPDF2.PdfReader(file)
                # Try to access pages to ensure it's readable
                num_pages = len(pdf.pages)
                return True
            except Exception:
                return False
    except ImportError:
        # If PyPDF2 is not installed, just check file size
        return os.path.getsize(pdf_path) > 100  # Assume valid if > 100 bytes

def is_file_locked(file_path):
    """Check if a file is locked (in use by another process)"""
    try:
        with open(file_path, 'r+b') as f:
            return False
    except IOError:
        return True

def get_optimal_batch_size():
    """Determine optimal batch size based on available system memory"""
    mem = psutil.virtual_memory()
    # Use a smaller batch size for systems with less RAM
    if mem.total < 8 * 1024 * 1024 * 1024:  # 8 GB
        return 5
    elif mem.total < 16 * 1024 * 1024 * 1024:  # 16 GB
        return 10
    else:
        return 15

def parse_args():
    parser = argparse.ArgumentParser(description='Convert Word documents to PDF and rename with Project IDs')
    parser.add_argument('--input', '-i', help='Input directory containing Word documents')
    parser.add_argument('--output', '-o', help='Output directory for PDF files')
    parser.add_argument('--rename', '-r', action='store_true', help='Rename files with project IDs', default=True)
    parser.add_argument('--no-rename', action='store_false', dest='rename', help="Don't rename files with project IDs")
    return parser.parse_args()

def normalize_path(path):
    """Ensure path is in a format Word can handle"""
    # Convert UNC paths to mapped drives if needed
    if path.startswith('\\\\'):
        # For UNC paths, consider using temporary local copies
        # or map a drive letter temporarily
        pass
    return os.path.abspath(path)

def extract_project_id_from_filename(filename):
    """
    Extract a project ID from a filename if present.
    Project IDs are in the format P followed by 6 digits,
    followed by either a hyphen or underscore (e.g., P123456- or P123456_).
    
    Args:
        filename: The filename to check
        
    Returns:
        The project ID if found, None otherwise
    """
    try:
        # Regular expression pattern for project ID in filename
        # Looks for P + 6 digits + (- or _)
        pattern = r'P\d{6}[-_]'
        
        # Find all matches in the filename
        matches = re.findall(pattern, filename)
        if matches:
            # Return the first match without the trailing - or _
            return matches[0][:-1]
                
        return None
        
    except Exception as e:
        logging.error(f"Error processing filename {filename}: {str(e)}")
        return None

def detect_language(pdf_path, pages_to_check=3):
    """
    Detect if a PDF document is primarily in English or not.
    
    Args:
        pdf_path: Path to the PDF file
        pages_to_check: Number of pages to analyze for language detection
        
    Returns:
        "EN" if English is detected, "NON" otherwise
    """
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # Limit the number of pages to check
            pages_to_check = min(len(reader.pages), pages_to_check)
            
            # Concatenate text from multiple pages for better detection
            all_text = ""
            for page_num in range(pages_to_check):
                page = reader.pages[page_num]
                text = page.extract_text()
                if text:
                    all_text += text
                    # Once we have a decent amount of text, we can stop
                    if len(all_text) > 1000:
                        break
            
            # If we have enough text to detect language
            if len(all_text) > 100:
                try:
                    lang = detect(all_text)
                    return "EN" if lang == "en" else "NON"
                except LangDetectException:
                    logging.warning(f"Could not detect language in {pdf_path}")
                    return "NON"  # Default to non-English if detection fails
            else:
                logging.warning(f"Not enough text for language detection in {pdf_path}")
                return "NON"  # Default to non-English if not enough text
                    
    except Exception as e:
        logging.error(f"Error detecting language in {pdf_path}: {str(e)}")
        return "NON"  # Default to non-English on error

if __name__ == "__main__":
    args = parse_args()
    if args.input and args.output:
        # TODO: Add command-line mode implementation
        pass
    else:
        sys.exit(convert_folder_to_pdf(rename_with_pid=True)) 