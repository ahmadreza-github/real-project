import os
import chardet
from docx import Document

folder_path = r'C:\Users\i7 11th\Desktop\qods files'

for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)

    if os.path.isfile(file_path):  # Ensure it's a file, not a folder

        # Check if it's a .docx file
        if filename.lower().endswith('.docx'):
            doc = Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs])

        else:
            # Detect encoding for text files
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)  # Read a chunk to detect encoding
                result = chardet.detect(raw_data)
                detected_encoding = result['encoding'] if result['encoding'] else 'utf-8'

            # Read with detected encoding
            with open(file_path, 'r', encoding=detected_encoding, errors='ignore') as f:
                content = f.read()

        # Check if there is content in the file
        if content.strip():
            print(f"There is something in {filename}")
        else:
            print(f"{filename} is empty.")


for x in