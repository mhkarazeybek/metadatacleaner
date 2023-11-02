import os
from PIL import Image
import PyPDF2
from docx import Document
from openpyxl import load_workbook
import argparse

def clear_metadata_image(image_path):
    try:
        image = Image.open(image_path)
        data = list(image.getdata())
        image_without_exif = Image.new(image.mode, image.size)
        image_without_exif.putdata(data)
        image_without_exif.save(image_path)
    except Exception as e:
        print(f"Error 1: {e}")

def clear_metadata_pdf(pdf_path):
    try:
        reader = PyPDF2.PdfFileReader(pdf_path)
        writer = PyPDF2.PdfFileWriter()

        for i in range(reader.getNumPages()):
            writer.addPage(reader.getPage(i))

        writer.encrypt('') 

        with open(pdf_path, 'wb') as new_file:
            writer.write(new_file)
    except Exception as e:
        print(f"Error 2: {e}")

def clear_metadata_word(doc_path):
    try:
        doc = Document(doc_path)
        doc.core_properties.authors = ""
        doc.core_properties.title = ""
        doc.core_properties.subject = ""
        doc.core_properties.creator = ""
        doc.core_properties.keywords = ""
        doc.core_properties.last_modified_by = ""
        doc.save(doc_path)
    except Exception as e:
        print(f"Error 3: {e}")

def clear_metadata_excel(excel_path):
    try:
        wb = load_workbook(excel_path)
        props = wb.properties
        props.creator = ""
        props.last_modified_by = ""
        props.title = ""
        props.subject = ""
        props.category = ""
        props.keywords = ""
        props.program = ""
        props.comments = ""
        props.company = ""
        wb.save(excel_path)
    except Exception as e:
        print(f"Error 4: {e}")

def clear_metadata_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                clear_metadata_image(file_path)
            elif file.lower().endswith('.pdf'):
                clear_metadata_pdf(file_path)
            elif file.lower().endswith('.docx'):
                clear_metadata_word(file_path)
            elif file.lower().endswith(('.xlsx', '.xls')):
                clear_metadata_excel(file_path)
            else:
                print(f"Unsupported: {file}")


def main():
    parser = argparse.ArgumentParser(description="Meta Data Cleaner")
    parser.add_argument("folder_path",type=str, help="Clean This Folder", default="test")
    args = parser.parse_args()

    if os.path.isdir(args.folder_path):
        clear_metadata_folder(args.folder_path)
    else:
        print("Folder Not Found")

    print("DONE")

if __name__ == "__main__":
    main()
