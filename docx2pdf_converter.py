#!/usr/bin/env python3
import os
import time
import argparse
import concurrent.futures
from pathlib import Path
import platform
import subprocess
from tqdm import tqdm

def convert_with_libreoffice(input_file, output_file=None):
    """Convert docx to PDF using LibreOffice in headless mode"""
    if output_file is None:
        output_file = str(Path(input_file).with_suffix('.pdf'))
    
    # Determine LibreOffice executable based on platform
    if platform.system() == "Windows":
        # Possible paths for LibreOffice on Windows
        libreoffice_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        soffice = next((p for p in libreoffice_paths if os.path.exists(p)), None)
    elif platform.system() == "Darwin":  # macOS
        soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:  # Linux and others
        soffice = "libreoffice"
    
    if not soffice:
        raise Exception("LibreOffice executable not found")
    
    output_dir = os.path.dirname(output_file) or '.'
    
    args = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_file
    ]
    
    process = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()
    
    if process.returncode != 0:
        raise Exception(f"LibreOffice conversion failed: {stderr.decode()}")
    # LibreOffice creates the PDF with the same name as the input file but with .pdf extension
    default_output = str(Path(os.path.join(output_dir, Path(input_file).stem)).with_suffix('.pdf'))
    
    # If the user specified a different output file name, rename the file
    if output_file != default_output and os.path.exists(default_output):
        os.rename(default_output, output_file)
    
    return output_file

def convert_with_word(input_file, output_file=None, retries=2, timeout=30):
    """Convert docx to PDF using Microsoft Word (Windows only)"""
    if output_file is None:
        output_file = str(Path(input_file).with_suffix('.pdf'))
    
    # Only import win32com if we're using this function
    import win32com.client
    import pythoncom
    import comtypes.client
    import time
    import threading
    
    # Remove the SIGALRM timeout as it doesn't work on Windows
    # Instead, use a threading-based timeout approach
    
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
                doc.Close()
                
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

def process_file(file_path, output_dir, use_word):
    """Process a single file conversion with error handling"""
    try:
        input_path = os.path.abspath(file_path)
        output_name = Path(file_path).stem + ".pdf"
        output_path = os.path.join(output_dir, output_name)
        
        # Ensure we're on Windows since Word is required
        if platform.system() != "Windows":
            raise Exception("Microsoft Word conversion requires Windows")
            
        # Convert using Word
        convert_with_word(input_path, output_path, retries=2)
        return (file_path, True, None)
    except Exception as e:
        return (file_path, False, str(e))

def main():
    parser = argparse.ArgumentParser(description='Convert DOCX files to PDF in parallel')
    parser.add_argument('input_dir', help='Directory containing DOCX files')
    parser.add_argument('--output-dir', '-o', help='Output directory for PDF files (default: same as input)')
    parser.add_argument('--workers', '-w', type=int, default=min(os.cpu_count(), 4),
                       help='Number of parallel worker processes')
    parser.add_argument('--max-retries', '-r', type=int, default=3,
                       help='Maximum number of retries for failed conversions')
    parser.add_argument('--batch-size', '-b', type=int, default=10,
                       help='Number of files to process before restarting Word (prevents memory leaks)')
    
    args = parser.parse_args()
    
    # Check if we're on Windows
    if platform.system() != "Windows":
        print("Error: This script requires Windows with Microsoft Word installed")
        return 1
    
    # Validate input directory
    if not os.path.isdir(args.input_dir):
        print(f"Error: {args.input_dir} is not a valid directory")
        return 1
    
    # Set up output directory
    output_dir = args.output_dir if args.output_dir else args.input_dir
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Find all .docx files
    docx_files = []
    for root, _, files in os.walk(args.input_dir):
        for file in files:
            if file.lower().endswith('.docx'):
                docx_files.append(os.path.join(root, file))
    
    if not docx_files:
        print(f"No .docx files found in {args.input_dir}")
        return 0
    
    print(f"Found {len(docx_files)} .docx files to convert")
    
    # Print configuration
    print("Using Microsoft Word for conversion")
    print(f"Using {args.workers} worker processes")
    
    # Initialize counters and timing
    start_time = time.time()
    successful = 0
    failed = 0
    
    # Process in smaller batches to prevent memory issues in corporate environments
    batch_size = args.batch_size
    for i in range(0, len(docx_files), batch_size):
        batch = docx_files[i:i+batch_size]
        
        print(f"\nProcessing batch {i//batch_size + 1} of {(len(docx_files) + batch_size - 1) // batch_size} ({len(batch)} files)")
        
        # Clean up any existing Word processes before each batch
        try:
            import subprocess
            subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], 
                          stdout=subprocess.DEVNULL, 
                          stderr=subprocess.DEVNULL)
            time.sleep(1)  # Give system time to close Word
        except:
            pass
                
        # Process the current batch
        with concurrent.futures.ProcessPoolExecutor(max_workers=args.workers) as executor:
            # Submit jobs for this batch
            future_to_file = {
                executor.submit(process_file, file, output_dir, True): file
                for file in batch
            }
            
            # Process results
            with tqdm(total=len(batch), unit="file") as pbar:
                for future in concurrent.futures.as_completed(future_to_file):
                    file_path, success, error = future.result()
                    if success:
                        successful += 1
                    else:
                        failed += 1
                        print(f"Error converting {file_path}: {error}")
                    pbar.update(1)
    
    # Report results
    elapsed_time = time.time() - start_time
    files_per_second = len(docx_files) / elapsed_time if elapsed_time > 0 else 0
    
    print(f"\nConversion complete in {elapsed_time:.2f} seconds ({files_per_second:.2f} files/sec)")
    print(f"Successfully converted: {successful}")
    print(f"Failed conversions: {failed}")
    
    # Add success rate report
    if successful + failed > 0:
        print(f"Success rate: {successful/(successful+failed)*100:.1f}%")
    else:
        print("Success rate: N/A (no files processed)")
    
    return 0

if __name__ == "__main__":
    exit(main()) 