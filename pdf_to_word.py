import os
import win32com.client

def pdf_to_word(file_path):
    # Create an instance of the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        # Open the PDF file and convert it to Word
        doc = word.Documents.Open(file_path, ReadOnly=True)
        doc.SaveAs(os.path.splitext(file_path)[0] + ".docx", FileFormat=16)
        
        # Close the document and delete the PDF file
        doc.Close()
        os.remove(file_path)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Close the Word application
        word.Quit()

def convert_folder_to_word(folder_path):
    # Recursively traverse the folder and its subfolders
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # Check if the file is a PDF document
            if os.path.splitext(file)[1] == ".pdf":
                # Construct the full path to the file and convert it to Word
                file_path = os.path.join(root, file)
                print(f"Converting {file_path} to Word...")
                pdf_to_word(file_path)

# Example usage: convert all PDF files in the "C:\MyFolder" folder and its subfolders to Word documents
convert_folder_to_word(r"H:\NA\dice ISO\new policies")
