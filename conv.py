# # first way that almost worked
# # def converter():
# #     from markitdown import MarkItDown
# #     import os
# #     import re

# #     def word_to_md(word_file_path, md_file_path):
# #         """Converts a Word file to a Markdown (.md) file."""
# #         try:
# #             md = MarkItDown() 
# #             result = md.convert(word_file_path)
# #             markdown_content = result.text_content

# #             # Optionally, you can perform additional Markdown cleanup or adjustments here
# #             # For example, ensure consistent heading levels, list formatting, etc.
# #             # However, the markitdown library does a fairly good job.

# #             with open(md_file_path, "w", encoding="utf-8") as md_file:
# #                 md_file.write(markdown_content.strip())  # Remove leading/trailing whitespace
# #             print(f"Conversion successful: {word_file_path} -> {md_file_path}")
# #         except Exception as e:
# #             print(f"Error during conversion: {e}")

# #     word_file =r"C:\Users\i7 11th\Desktop\fold\third file.docx"  # Replace with your Word file path
# #     md_file = "output.md"
# #     word_to_md(word_file, md_file)

# # converter() 














# import os
# import pandas as pd
# from openai import OpenAI
# from openpyxl import Workbook, load_workbook
# from openpyxl.styles import Alignment
# from docx import Document
# import sys
# word_path = r'C:\Users\i7 11th\Desktop\New Microsoft Word Document (2).docx'

# if len(sys.argv) == 2: 
#     command = sys.argv[1]
# def llm_response():
#     source_word = r"C:\Users\i7 11th\Desktop\qods files\فرم_شناسنامه_تحلیلی_نهایی_شده_docx_زیارت.docx"
#     client = OpenAI(
#             base_url="https://openrouter.ai/api/v1",
#             api_key="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",
#         )

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
#                     "content": ""
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
#     print(response_text)
#     doc = Document()
#     doc.add_paragraph(response_text)
#     doc.save(word_path)

# if command == "main":
#     llm_response()
# elif command == 'path':
#     os.startfile(r"C:\Users\i7 11th\Desktop\New Microsoft Word Document (2).docx")
# elif command == "source":
#     os.startfile(r"C:\Users\i7 11th\Desktop\qods files\فرم_شناسنامه_تحلیلی_نهایی_شده_docx_زیارت.docx")
# else:
#     quit()





from groq import Groq

# Initialize the client with the API key
client = Groq(api_key="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")

# Make a request to the Groq API
completion = client.chat.completions.create(
    model="qwen-2.5-32b",
    messages=[{"role": "user", "content": "Tell me what you are and the company that created you"}],
    temperature=0.6,
    max_tokens=4096,  # Corrected from max_completion_tokens
    top_p=0.95,
    stream=True  # Removed 'stop=None', as it's not required
)

# Print the response
for chunk in completion:
    print(chunk.choices[0].delta.content, end="", flush=True)

