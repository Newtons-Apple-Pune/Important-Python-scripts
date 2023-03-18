import os
import win32com.client

def word_to_pdf(file_path):
    # Create an instance of the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        # Open the Word document and convert it to PDF
        doc = word.Documents.Open(file_path)
        pdf_path = os.path.splitext(file_path)[0] + ".pdf"
        doc.SaveAs(pdf_path, FileFormat=17)
        
        # Close the document and delete the Word file
        doc.Close()
        os.remove(file_path)
        
        return pdf_path
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Close the Word application
        word.Quit()

def convert_folder_to_pdf(folder_path):
    # Recursively traverse the folder and its subfolders
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # Check if the file is a Word document
            if os.path.splitext(file)[1] in (".doc", ".docx"):
                # Construct the full path to the file and convert it to PDF
                file_path = os.path.join(root, file)
                print(f"Converting {file_path} to PDF...")
                pdf_path = word_to_pdf(file_path)
                
                if pdf_path:
                    # Rename the PDF file to the original name of the Word file
                    os.rename(pdf_path, os.path.splitext(file_path)[0] + ".pdf")

# Example usage: convert all Word documents in the "C:\MyFolder" folder and its subfolders to PDF and replace the original Word documents with their PDF forms
convert_folder_to_pdf(r"H:\NA\dice ISO\auxilary docs")
