import os
import re
import PyPDF2
from pathlib import Path
from typing import Optional, Tuple
from docx2pdf import convert

def convert_word_to_pdf(word_path: str) -> Optional[str]:
    """
    Convert a Word document (.doc or .docx) to PDF format.
    
    Args:
        word_path: Path to the Word document
        
    Returns:
        Path to the converted PDF file if successful, None otherwise
    """
    try:
        pdf_path = str(Path(word_path).with_suffix('.pdf'))
        convert(word_path, pdf_path)
        return pdf_path
    except Exception as e:
        print(f"Error converting {word_path} to PDF: {str(e)}")
        return None

def extract_project_id(pdf_path: str, max_pages: int = 10) -> Optional[str]:
    """
    Extract the first occurrence of a World Bank project ID from a PDF file.
    Project IDs are in the format P followed by 6 digits (e.g., P123456).
    
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
            
            # Regular expression pattern for project ID
            pattern = r'P\d{6}'
            
            # Search through pages
            for page_num in range(pages_to_search):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                # Find all matches in the page
                matches = re.findall(pattern, text)
                if matches:
                    return matches[0]  # Return the first match
                    
        return None
        
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return None

def process_and_rename_document(doc_path: Path, folder: Path) -> bool:
    """
    Process a document (PDF or Word) and rename it based on project ID.
    For Word documents, converts to PDF and removes the original.
    
    Args:
        doc_path: Path to the document
        folder: Path to the parent folder
        
    Returns:
        True if successfully processed and renamed, False otherwise
    """
    pdf_path = str(doc_path)
    is_word = doc_path.suffix.lower() in ['.doc', '.docx']
    converted_pdf = None
    
    try:
        # Convert Word documents to PDF
        if is_word:
            print(f"Converting Word document: {doc_path.name}")
            converted_pdf = convert_word_to_pdf(str(doc_path))
            if not converted_pdf:
                return False
            pdf_path = converted_pdf
        
        # Extract project ID
        project_id = extract_project_id(pdf_path)
        
        if project_id:
            # Create new filename with project ID (always .pdf)
            new_filename = f"{project_id}.pdf"
            new_path = folder / new_filename
            
            # Handle duplicate filenames
            counter = 1
            while new_path.exists():
                new_filename = f"{project_id}_{counter}.pdf"
                new_path = folder / new_filename
                counter += 1
            
            if is_word:
                # For Word documents, move the converted PDF to the new name
                Path(converted_pdf).rename(new_path)
                # Delete the original Word document
                doc_path.unlink()
                print(f"Converted and renamed: {doc_path.name} -> {new_filename}")
            else:
                # For PDFs, just rename
                doc_path.rename(new_path)
                print(f"Renamed: {doc_path.name} -> {new_filename}")
            
            return True
        else:
            print(f"No project ID found in: {doc_path.name}")
            # If this was a Word doc conversion with no project ID, clean up the PDF
            if converted_pdf:
                Path(converted_pdf).unlink()
            return False
            
    except Exception as e:
        print(f"Error processing {doc_path.name}: {str(e)}")
        # Clean up converted PDF if it exists
        if converted_pdf:
            try:
                Path(converted_pdf).unlink()
            except:
                pass
        return False

def rename_documents_with_project_ids(folder_path: str) -> Tuple[int, int]:
    """
    Process all documents in the specified folder:
    - Rename PDF files based on project IDs
    - Convert Word files to PDF, rename based on project IDs, and remove originals
    
    Args:
        folder_path: Path to the folder containing documents
        
    Returns:
        Tuple of (number of successfully processed files, total number of files)
    """
    folder = Path(folder_path)
    if not folder.exists():
        raise ValueError(f"Folder not found: {folder_path}")
    
    successful_processes = 0
    total_docs = 0
    
    # Process all PDF and Word files in the folder
    for pattern in ['*.pdf', '*.doc', '*.docx']:
        for doc_file in folder.glob(pattern):
            total_docs += 1
            print(f"\nProcessing: {doc_file.name}")
            
            if process_and_rename_document(doc_file, folder):
                successful_processes += 1
    
    return successful_processes, total_docs

def main():
    """
    Main function to run the document renaming script.
    """
    # Get folder path from user and remove any quotes
    folder_path = input("Enter the folder path containing PDF and Word files: ").strip().strip('"\'').strip()
    
    try:
        successful, total = rename_documents_with_project_ids(folder_path)
        print(f"\nSummary:")
        print(f"Total documents processed: {total}")
        print(f"Successfully processed and renamed: {successful}")
        print(f"Failed to process: {total - successful}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()