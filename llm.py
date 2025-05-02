# import os
# import pandas as pd
# from openai import OpenAI
# from openpyxl import Workbook, load_workbook
# from openpyxl.styles import Alignment
# def for_multiple_choice():
#     # Ensure the required file exists
#     file_path = "outpur.md" 
#     if not os.path.exists(file_path):
#         print("❌ Error: File 'output.md' not found!")
#         exit()

#     # Read the content of the markdown file
#     with open(file_path, "r", encoding="utf-8") as md_file:
#         file_content = md_file.read().strip()

#     # OpenAI API call
#     client = OpenAI(
#         base_url="https://openrouter.ai/api/v1",
#         api_key="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",
#     )

#     try:
#         completion = client.chat.completions.create(
#             extra_headers={
#                 "HTTP-Referer": "<YOUR_SITE_URL>",
#                 "X-Title": "<YOUR_SITE_NAME>",
#             },
#             model="deepseek/deepseek-r1-distill-llama-70b:free",
#             messages=[
#                 {
#                     "role": "user",
#                     "content": f"""
#                     You are an AI specialized in text analysis and structured data extraction. Carefully analyze the provided text, focusing only on multiple-choice questions that contain percentage-based responses. Your task is to extract and structure the relevant information in a precise and structured format.

#                     ### **Extraction Guidelines:**
#                     1. **Extract only relevant multiple-choice responses** that include percentages.
#                     2. **Remove all symbols** (such as hollow or filled squares/circles) and keep the extracted text clear and concise.
#                     3. **Ensure the output is in CSV format** (comma-separated values) and properly structured.
#                     4. **Column Headings for CSV File:**
#                     - "مورد" (Case): The exact item described in the question.
#                     - "درصد" (Percentage): The associated percentage.

#                     5. **Handling Unclear or Ambiguous Cases:**
#                     - If a correct percentage cannot be determined, return **"مشکل رخ داد" (Problem occurred)** in the percentage field.
#                     - If the data is ambiguous or unclear, return **"داده نامشخص" (Unknown data)**.

#                     ### **Output Formatting:**
#                     - The response must be encoded in **UTF-8** to correctly handle Persian text.
#                     - Each entry should be **on a new line** following the CSV structure.
#                     - **No extra elaboration is allowed**—only structured responses as required.

#                     ### **Example Input (Extracted from Survey Text):**
#                     ### **Expected Output (CSV Format, UTF-8 Encoding)**
#                     **Final Instructions:**
#                     - Ensure every extracted value follows this format.
#                     - **Ignore irrelevant text** that does not match the required structure.
#                     - Provide only the requested structured data **without extra details**.
#                     here is the survey text:**
                    
#                     {file_content}
#                     """
#                 }
#             ]
#         )
#     except Exception as e:
#         print(f"❌ OpenAI API Request Failed: {e}")
#         exit()

#     # Ensure response is valid
#     if not completion.choices or not completion.choices[0].message:
#         print("❌ Error: No valid response received from OpenAI API")
#         exit()

#     # Extract and clean the response
#     response_text = completion.choices[0].message.content.strip()
#     print("✅ Raw Output from LLM:\n", response_text)  # Debugging output

#     # Convert extracted text into a structured format
#     try:
#         data = []
#         for line in response_text.split("\n"):
#             parts = line.split(",", 1)  # Split into question & answer
#             if len(parts) == 2:
#                 data.append([parts[0].strip(), parts[1].strip()])

#         # Ensure data exists
#         if not data:
#             print("❌ Error: No valid question-answer pairs found in AI response.")
#             exit()

#         # Save to an Excel file with RTL formatting
#         output_excel_file = r"C:\Users\i7 11th\Desktop\Excel result.xlsx"
#         wb = Workbook()
#         ws = wb.active

#         # Write headers
#         headers = ["سوال", "پاسخ صحیح"]
#         for col_idx, header in enumerate(headers, start=1):
#             cell = ws.cell(row=1, column=col_idx, value=header)
#             cell.alignment = Alignment(horizontal="right")  # Align headers to right

#         # Write data with RTL formatting
#         for row_idx, row in enumerate(data, start=2):
#             for col_idx, value in enumerate(row, start=1):
#                 cell = ws.cell(row=row_idx, column=col_idx, value=value)
#                 cell.alignment = Alignment(horizontal="right")  # Align content to right

#         # Adjust column width dynamically
#         for col in ws.columns:
#             max_length = max(len(str(cell.value)) for cell in col if cell.value)  # Get max length
#             ws.column_dimensions[col[0].column_letter].width = max_length + 2  # Adjust width

#         # Save the file
#         wb.save(output_excel_file)

#         print(f"✅ Data successfully saved to {output_excel_file} with Persian RTL formatting!")

#     except Exception as e:
#         print(f"❌ Error processing or saving to Excel: {e}")

# import os
# import openpyxl
# from openai import OpenAI
# from openpyxl.styles import Alignment

# def for_tables():
#     """Extracts and formats Persian survey tables into a structured Excel file."""

#     # ✅ Ensure the required file exists
#     file_path = "outpur.md"
#     if not os.path.exists(file_path):
#         print("❌ Error: File 'output.md' not found!")
#         return

#     # ✅ Read the content of the markdown file
#     with open(file_path, "r", encoding="utf-8") as md_file:
#         file_content = md_file.read().strip()

#     # ✅ OpenAI API Call
#     client = OpenAI(
#         base_url="https://openrouter.ai/api/v1",
#         api_key="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",
#     )

#     try:
#         completion = client.chat.completions.create(
#             extra_headers={
#                 "HTTP-Referer": "<YOUR_SITE_URL>",
#                 "X-Title": "<YOUR_SITE_NAME>",
#             },
#             model="deepseek/deepseek-r1-distill-llama-70b:free",
#             messages=[
#                 {
#                     "role": "user",
#                     "content": f"""
#                     📌 **Task:** You are an AI specialized in text analysis and structured data extraction.
#                     **Extract & summarize information from tables** in the provided Persian survey text.

#                     🔹 **Instructions:**
#                     - **Extract only the meaningful content** from tables (ignore structure).
#                     - **Summarize long explanations** while keeping their meaning intact.
#                     - **Format as CSV** (Persian right-to-left format).
#                     - **Do not include any symbols** (e.g., hollow/filled squares, circles).
#                     - **Ensure proper UTF-8 encoding** for Persian text.
#                     - **Ensure all your answers are correct and in Farsi.**
#                     - ** as your answer is going to be automatically saved as a csv file, say nothing special more than necessary information.
#                     - ** If there are names and and information of people, include their, title, name, surname and all the information presented about them. **-
#                     - ** Try to be flexible and make every necessary change for a understandable ouput, you can include every necessary heading for the presented information.
#                     🔹 **Error Handling:**
#                     - If the correct answer is unclear, return: "مشکل رخ داد" (Problem occurred).
#                     - If data is ambiguous or missing, return: "داده نامشخص" (Unknown data).

#                     🔹 **Expected Output (CSV Format in Persian):**
#                     ```
#                     شاخص,مقدار
#                     جمعیت محله,۵۰۰۰ نفر (رشد ۲٪ در سال گذشته)
#                     نوع منازل,۶۰٪ آپارتمانی، ۴۰٪ ویلایی (وضعیت کلی مناسب است)
#                     وضعیت معابر,۳۰٪ خاکی، ۷۰٪ آسفالت (برخی کوچه‌ها نیازمند ترمیم هستند)
#                     ```

#                     {file_content}
#                     """
#                 }
#             ]
#         )
#     except Exception as e:
#         print(f"❌ OpenAI API Request Failed: {e}")
#         return

#     # ✅ Ensure response is valid
#     if not completion.choices or not completion.choices[0].message:
#         print("❌ Error: No valid response received from OpenAI API")
#         return

#     # ✅ Extract and clean the response
#     response_text = completion.choices[0].message.content.strip()
#     print("✅ Raw Output from LLM:\n", response_text)  # Debugging output

#     # ✅ Convert extracted text into a structured format for Excel
#     try:
#         data = []
#         for line in response_text.split("\n"):
#             parts = line.split(",", 1)  # Split into two columns
#             if len(parts) == 2:
#                 data.append([parts[0].strip(), parts[1].strip()])

#         # ✅ Ensure data exists
#         if not data:
#             print("❌ Error: No valid table data extracted from AI response.")
#             return

#         # ✅ Save to an Excel file with RTL formatting
#         output_excel_file = r"C:\Users\i7 11th\Desktop\Excel result.xlsx"
#         wb = openpyxl.Workbook()
#         ws = wb.active

#         # ✅ Write headers
#         headers = ["شاخص", "مقدار"]
#         for col_idx, header in enumerate(headers, start=1):
#             cell = ws.cell(row=1, column=col_idx, value=header)
#             cell.alignment = Alignment(horizontal="right")  # Align headers to right

#         # ✅ Write extracted data
#         for row_idx, row in enumerate(data, start=2):
#             for col_idx, value in enumerate(row, start=1):
#                 cell = ws.cell(row=row_idx, column=col_idx, value=value)
#                 cell.alignment = Alignment(horizontal="right")  # Align content to right

#         # ✅ Adjust column width dynamically
#         for col in ws.columns:
#             max_length = max(len(str(cell.value)) for cell in col if cell.value)  # Get max length
#             ws.column_dimensions[col[0].column_letter].width = max_length + 2  # Adjust width

#         # ✅ Save the file
#         wb.save(output_excel_file)

#         print(f"✅ Data successfully saved to {output_excel_file} with Persian RTL formatting!")

#     except Exception as e:
#         print(f"❌ Error processing or saving to Excel: {e}")

# # ✅ Run the function
# # for_tables()

# def for_descritive_questions():
#     """Extracts and formats Persian survey tables into a structured Excel file."""

#     # ✅ Ensure the required file exists
#     file_path = "output.md"
#     if not os.path.exists(file_path):
#         print("❌ Error: File 'output.md' not found!")
#         return

#     # ✅ Read the content of the markdown file
#     with open(file_path, "r", encoding="utf-8") as md_file:
#         file_content = md_file.read().strip()

#     # ✅ OpenAI API Call
#     client = OpenAI(
#         base_url="https://openrouter.ai/api/v1",
#         api_key="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",
#     )

#     try:
#         completion = client.chat.completions.create(
#             extra_headers={
#                 "HTTP-Referer": "<YOUR_SITE_URL>",
#                 "X-Title": "<YOUR_SITE_NAME>",
#             },
#             model="deepseek/deepseek-r1-distill-llama-70b:free",
#             messages=[
#                 {
#                     "role": "user",
#                     "content": f"""
#                     Task: You are an AI specializing in text analysis, structured data extraction, and summarization. Your objective is to extract meaningful content from descriptive questions and answers, ensuring clarity, accuracy, and proper formatting.

#                         🔹 Instructions:
#                         ✅ Content Extraction & Summarization:

#                         Extract only meaningful content from the provided questions and answers. Ignore unnecessary formatting (e.g., bullet points, numbering, decorative symbols).
#                         Summarize lengthy explanations while preserving their original meaning and key details.
#                         Maintain Persian (RTL) text format and ensure correct UTF-8 encoding.
#                         If a person’s details are mentioned, include all available information (e.g., title, first name, last name, and any additional details).
#                         If needed, make adjustments to improve clarity while preserving all necessary information.
#                         ✅ Output Structure & Formatting:

#                         Format the output as a structured CSV file suitable for direct processing.
#                         Include two primary columns in the CSV:
#                         "سوال" (Question) – Extracted and, if necessary, shortened version of the question, enclosed in quotes.
#                         "پاسخ" (Answer) – Extracted and, if necessary, summarized version of the answer, enclosed in quotes.
#                         Ensure readability and consistency in formatting.
#                         ✅ Error Handling:

#                         If the correct answer is unclear or incomplete, return: "مشکل رخ داد" (Problem occurred).
#                         If data is ambiguous or missing, return: "داده نامشخص" (Unknown data).

#                     {file_content}
#                     """
#                 }
#             ]
#         )
#     except Exception as e:
#         print(f"❌ OpenAI API Request Failed: {e}")
#         return

#     # ✅ Ensure response is valid
#     if not completion.choices or not completion.choices[0].message:
#         print("❌ Error: No valid response received from OpenAI API")
#         return

#     # ✅ Extract and clean the response
#     response_text = completion.choices[0].message.content.strip()
#     print("✅ Raw Output from LLM:\n", response_text)  # Debugging output

#     # ✅ Convert extracted text into a structured format for Excel
#     try:
#         data = []
#         for line in response_text.split("\n"):
#             parts = line.split(",", 1)  # Split into two columns
#             if len(parts) == 2:
#                 data.append([parts[0].strip(), parts[1].strip()])

#         # ✅ Ensure data exists
#         if not data:
#             print("❌ Error: No valid table data extracted from AI response.")
#             return

#         # ✅ Save to an Excel file with RTL formatting
#         output_excel_file = r"C:\Users\i7 11th\Desktop\Excel result.xlsx"
#         wb = openpyxl.Workbook()
#         ws = wb.active

#         # ✅ Write headers
#         headers = ["شاخص", "مقدار"]
#         for col_idx, header in enumerate(headers, start=1):
#             cell = ws.cell(row=1, column=col_idx, value=header)
#             cell.alignment = Alignment(horizontal="right")  # Align headers to right

#         # ✅ Write extracted data
#         for row_idx, row in enumerate(data, start=2):
#             for col_idx, value in enumerate(row, start=1):
#                 cell = ws.cell(row=row_idx, column=col_idx, value=value)
#                 cell.alignment = Alignment(horizontal="right")  # Align content to right

#         # ✅ Adjust column width dynamically
#         for col in ws.columns:
#             max_length = max(len(str(cell.value)) for cell in col if cell.value)  # Get max length
#             ws.column_dimensions[col[0].column_letter].width = max_length + 2  # Adjust width

#         # ✅ Save the file
#         wb.save(output_excel_file)

#         print(f"✅ Data successfully saved to {output_excel_file} with Persian RTL formatting!")

#     except Exception as e:
#         print(f"❌ Error processing or saving to Excel: {e}")
# for_descritive_questions()






















# let's test 2:
# import os
# import docx
# from docx import Document

# def extract_tables(doc):
#     """Extract tables from the Word document and format them as Markdown."""
#     md_content = []
#     for table in doc.tables:
#         md_content.append("| " + " | ".join([cell.text.strip() for cell in table.rows[0].cells]) + " |")
#         md_content.append("| " + " | ".join(["---" for _ in table.rows[0].cells]) + " |")
#         for row in table.rows[1:]:
#             md_content.append("| " + " | ".join([cell.text.strip() for cell in row.cells]) + " |")
#         md_content.append("\n")
#     return "\n".join(md_content)

# def extract_multiple_choice(doc):
#     """Extract multiple-choice questions with symbols and percentages."""
#     md_content = []
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if any(symbol in text for symbol in ["▪", "▫", "■", "□", "✔", "✖"]):
#             md_content.append(f"- {text}")
#     return "\n".join(md_content)

# def extract_descriptive(doc):
#     """Extract descriptive questions and answers."""
#     md_content = []
#     current_question = ""
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if text.endswith("؟"):  # Checking for Persian question mark
#             if current_question:
#                 md_content.append(f"### {current_question}\n")
#             current_question = text
#         elif text:
#             md_content.append(text)
#     if current_question:
#         md_content.append(f"### {current_question}\n")
#     return "\n".join(md_content)

# def convert_docx_to_md(file_path, output_md):
#     """Process the .docx file and create an integrated .md file."""
#     doc = Document(file_path)
#     md_content = []
    
#     # Extract and format data
#     md_content.append("## Tables\n" + extract_tables(doc))
#     md_content.append("## Multiple Choice Questions\n" + extract_multiple_choice(doc))
#     md_content.append("## Descriptive Questions\n" + extract_descriptive(doc))
    
#     # Write to Markdown file
#     with open(output_md, "w", encoding="utf-8") as md_file:
#         md_file.write("\n".join(md_content))
    
#     print(f"✅ Markdown file created: {output_md}")

# # Example Usage
# convert_docx_to_md(r"C:\Users\i7 11th\Desktop\فرم شناسنامه تحلیلی نهایی شده.docx", "pro_result.md")

# let's test 3:
# import os
# import docx
# from docx import Document

# def extract_tables(doc):
#     """Extract tables from the Word document and format them as Markdown."""
#     md_content = []
#     for table in doc.tables:
#         md_content.append("| " + " | ".join([cell.text.strip() for cell in table.rows[0].cells]) + " |")
#         md_content.append("| " + " | ".join(["---" for _ in table.rows[0].cells]) + " |")
#         for row in table.rows[1:]:
#             md_content.append("| " + " | ".join([cell.text.strip() for cell in row.cells]) + " |")
#         md_content.append("\n")
#     return "\n".join(md_content)

# def extract_multiple_choice(doc):
#     """Extract multiple-choice questions with symbols and percentages, preserving all symbols."""
#     md_content = []
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if any(symbol in text for symbol in ["▪", "▫", "■", "□", "✔", "✖", "○", "●", "✓", "✗"]):
#             md_content.append(f"- {text}")
#     return "\n".join(md_content)

# def extract_descriptive(doc):
#     """Extract descriptive questions and answers."""
#     md_content = []
#     current_question = ""
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if text.endswith("؟"):  # Checking for Persian question mark
#             if current_question:
#                 md_content.append(f"### {current_question}\n")
#             current_question = text
#         elif text:
#             md_content.append(text)
#     if current_question:
#         md_content.append(f"### {current_question}\n")
#     return "\n".join(md_content)

# def convert_docx_to_md(file_path, output_md):
#     """Process the .docx file and create an integrated .md file, ensuring symbols are preserved."""
#     doc = Document(file_path)
#     md_content = []
    
#     # Extract and format data
#     md_content.append("## Tables\n" + extract_tables(doc))
#     md_content.append("## Multiple Choice Questions\n" + extract_multiple_choice(doc))
#     md_content.append("## Descriptive Questions\n" + extract_descriptive(doc))
    
#     # Write to Markdown file
#     with open(output_md, "w", encoding="utf-8") as md_file:
#         md_file.write("\n".join(md_content))
    
#     print(f"✅ Markdown file created: {output_md}")

# # Example Usage
# convert_docx_to_md(r"C:\Users\i7 11th\Desktop\فرم شناسنامه تحلیلی نهایی شده.docx", "pro_result.md")

# let's test 4:
# import os
# import docx
# from docx import Document
# import re

# def extract_tables(doc):
#     """Extract tables from the Word document and format them as Markdown."""
#     md_content = []
#     for table in doc.tables:
#         md_content.append("| " + " | ".join([cell.text.strip() for cell in table.rows[0].cells]) + " |")
#         md_content.append("| " + " | ".join(["---" for _ in table.rows[0].cells]) + " |")
#         for row in table.rows[1:]:
#             md_content.append("| " + " | ".join([cell.text.strip() for cell in row.cells]) + " |")
#         md_content.append("\n")
#     return "\n".join(md_content)

# def extract_multiple_choice(doc):
#     """Extract multiple-choice questions dynamically by detecting unique symbols."""
#     md_content = []
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         # Detect special symbols dynamically
#         match = re.search(r'([^\w\s]+)', text)  # Match non-alphanumeric, non-space characters
#         if match:
#             md_content.append(f"- {text}")
#     return "\n".join(md_content)

# def extract_descriptive(doc):
#     """Extract descriptive questions and answers."""
#     md_content = []
#     current_question = ""
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if text.endswith("؟"):  # Checking for Persian question mark
#             if current_question:
#                 md_content.append(f"### {current_question}\n")
#             current_question = text
#         elif text:
#             md_content.append(text)
#     if current_question:
#         md_content.append(f"### {current_question}\n")
#     return "\n".join(md_content)

# def convert_docx_to_md(file_path, output_md):
#     """Process the .docx file and create an integrated .md file, ensuring symbols are preserved dynamically."""
#     doc = Document(file_path)
#     md_content = []
    
#     # Extract and format data
#     md_content.append("## Tables\n" + extract_tables(doc))
#     md_content.append("## Multiple Choice Questions\n" + extract_multiple_choice(doc))
#     md_content.append("## Descriptive Questions\n" + extract_descriptive(doc))
    
#     # Write to Markdown file
#     with open(output_md, "w", encoding="utf-8") as md_file:
#         md_file.write("\n".join(md_content))
    
#     print(f"✅ Markdown file created: {output_md}")

# # Example Usage
# convert_docx_to_md(r"C:\Users\i7 11th\Desktop\فرم شناسنامه تحلیلی نهایی شده.docx", "pro_result.md")

# let's test 5:
# import os
# import docx
# from docx import Document
# import re

# def extract_tables(doc):
#     """Extract tables from the Word document and format them as Markdown."""
#     md_content = []
#     for table in doc.tables:
#         md_content.append("| " + " | ".join([cell.text.strip() for cell in table.rows[0].cells]) + " |")
#         md_content.append("| " + " | ".join(["---" for _ in table.rows[0].cells]) + " |")
#         for row in table.rows[1:]:
#             md_content.append("| " + " | ".join([cell.text.strip() for cell in row.cells]) + " |")
#         md_content.append("\n")
#     return "\n".join(md_content)

# def extract_multiple_choice(doc):
#     """Extract multiple-choice questions while preserving special symbols."""
#     md_content = []
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if any(ord(char) > 127 for char in text):  # Detect non-ASCII symbols
#             md_content.append(f"- {text}")
#     return "\n".join(md_content)

# def extract_descriptive(doc):
#     """Extract descriptive questions and answers."""
#     md_content = []
#     current_question = ""
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if text.endswith("؟"):  # Checking for Persian question mark
#             if current_question:
#                 md_content.append(f"### {current_question}\n")
#             current_question = text
#         elif text:
#             md_content.append(text)
#     if current_question:
#         md_content.append(f"### {current_question}\n")
#     return "\n".join(md_content)

# def convert_docx_to_md(file_path, output_md):
#     """Process the .docx file and create an integrated .md file, ensuring symbols are preserved."""
#     doc = Document(file_path)
#     md_content = []
    
#     # Extract and format data
#     md_content.append("## Tables\n" + extract_tables(doc))
#     md_content.append("## Multiple Choice Questions\n" + extract_multiple_choice(doc))
#     md_content.append("## Descriptive Questions\n" + extract_descriptive(doc))
    
#     # Write to Markdown file
#     with open(output_md, "w", encoding="utf-8") as md_file:
#         md_file.write("\n".join(md_content))
    
#     print(f"✅ Markdown file created: {output_md}")

# # Example Usage
# convert_docx_to_md(r"C:\Users\i7 11th\Desktop\فرم شناسنامه تحلیلی نهایی شده.docx", "pro_result.md")
# let's test 6:
# import os
# from docx import Document
# import unicodedata

# def contains_special_symbol(text):
#     """
#     Check if the given text contains any special symbol by looking up
#     each character's Unicode name and matching against a list of keywords.
#     """
#     keywords = ["BULLET", "DINGBAT", "CHECK", "BOX", "CIRCLE", "SQUARE", "STAR", "SYMBOL"]
#     for char in text:
#         try:
#             name = unicodedata.name(char)
#             if any(keyword in name for keyword in keywords):
#                 return True
#         except ValueError:
#             # If the character has no name, skip it
#             continue
#     return False

# def extract_tables(doc):
#     """Extract tables from the Word document and format them as Markdown."""
#     md_content = []
#     for table in doc.tables:
#         # Convert first row to header
#         header = "| " + " | ".join([cell.text.strip() for cell in table.rows[0].cells]) + " |"
#         separator = "| " + " | ".join(["---" for _ in table.rows[0].cells]) + " |"
#         md_content.append(header)
#         md_content.append(separator)
#         # Convert remaining rows
#         for row in table.rows[1:]:
#             row_text = "| " + " | ".join([cell.text.strip() for cell in row.cells]) + " |"
#             md_content.append(row_text)
#         md_content.append("\n")
#     return "\n".join(md_content)

# def extract_multiple_choice(doc):
#     """
#     Extract multiple-choice questions by scanning each paragraph.
#     If a paragraph contains any special symbols (as determined by its Unicode name),
#     it will be flagged and added to the multiple-choice section.
#     """
#     md_content = []
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if text and contains_special_symbol(text):
#             md_content.append(f"- {text}")
#     return "\n".join(md_content)

# def extract_descriptive(doc):
#     """Extract descriptive questions and answers."""
#     md_content = []
#     current_question = ""
#     for para in doc.paragraphs:
#         text = para.text.strip()
#         if text.endswith("؟"):  # Check for Persian question mark
#             if current_question:
#                 md_content.append(f"### {current_question}\n")
#             current_question = text
#         elif text:
#             md_content.append(text)
#     if current_question:
#         md_content.append(f"### {current_question}\n")
#     return "\n".join(md_content)

# def convert_docx_to_md(file_path, output_md):
#     """Process the .docx file and create a single integrated Markdown (.md) file."""
#     doc = Document(file_path)
#     md_sections = []
    
#     # Extract each section and add a header to the Markdown file.
#     md_sections.append("## Tables\n" + extract_tables(doc))
#     md_sections.append("## Multiple Choice Questions\n" + extract_multiple_choice(doc))
#     md_sections.append("## Descriptive Questions\n" + extract_descriptive(doc))
    
#     # Write combined content to the output Markdown file.
#     with open(output_md, "w", encoding="utf-8") as md_file:
#         md_file.write("\n".join(md_sections))
    
#     print(f"✅ Markdown file created: {output_md}")

# # Example Usagez
# convert_docx_to_md(r"C:\Users\i7 11th\Desktop\فرم شناسنامه تحلیلی نهایی شده.docx", "pro_result.md")
