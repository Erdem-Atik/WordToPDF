import os
import win32com.client as win32

# path to the folder containing the Word files
folder_path = r"D:\Coding\CreateWordFromExcel\output_folder"

# create the Word application object
word = win32.gencache.EnsureDispatch("Word.Application")

# loop through all the Word files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith(".doc") or file_name.endswith(".docx"):
        try:
            # open the Word file using Microsoft Word
            doc = word.Documents.Open(os.path.join(folder_path, file_name))

            # save the Word file as a PDF
            pdf_file_name = os.path.splitext(file_name)[0] + ".pdf"
            pdf_file_path = r"D:\Coding\CreateWordFromExcel\pdf_folder"
            doc.SaveAs(os.path.join(pdf_file_path, pdf_file_name), FileFormat=win32.constants.wdFormatPDF)

            # close the Word file
            doc.Close()
        except Exception as e:
            print(f"Failed to convert {file_name}: {str(e)}")

# quit the Word application
word.Quit()
