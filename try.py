import docx
import re
from openai import OpenAI

def extract_mcq_questions(docx_path):
    """
    Extracts multiple-choice questions and their corresponding options from a .docx file.
    """
    doc = docx.Document(docx_path)
    questions = []
    current_question = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue  # Skip empty lines

        # Detect question (assume it starts with a number followed by a dot)
        if re.match(r'^\d+\.', text):
            if current_question:
                questions.append(current_question)
            current_question = {"question": text, "options": []}

        # Detect options (assume they start with A), B), C), etc.)
        elif re.match(r'^[A-H]\)', text):
            if current_question:
                checkmarks = detect_checkmarks(para)
                current_question["options"].append({"marker": text[0], "text": text[2:].strip(), "checks": checkmarks})

    if current_question:
        questions.append(current_question)

    return questions

def detect_checkmarks(para):
    """
    Detects the number of checkmarks (✓) near an answer option.
    """
    checkmark_count = para.text.count("✓")
    return checkmark_count

def get_llm_response(prompt, client):
    """
    Sends a prompt to OpenAI's GPT model and returns the response.
    """
    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    return response.choices[0].message.content.strip()

def process_mcq_questions(docx_path, client):
    """
    Determines the selected answer for each MCQ using extracted checkmarks.
    """
    questions = extract_mcq_questions(docx_path)
    qa_pairs = []
    
    for q in questions:
        question_text = q["question"]
        options = q["options"]

        if not options:  # ✅ Prevent crash if options are missing
            print(f"⚠ Warning: No options found for question: {question_text}")
            continue  

        max_checks = max(opt["checks"] for opt in options) if options else 0
        selected_options = [opt["text"] for opt in options if opt["checks"] == max_checks]

        # Use LLM if multiple answers have the same check count
        if len(selected_options) > 1:
            options_text = "\n".join([f"{opt['marker']}) {opt['text']} [✓ count: {opt['checks']}]" for opt in options])
            
            prompt = f"""You are an AI specialized in interpreting multiple-choice questions in Persian.
The following is a multiple-choice question with its options.
Question:
{question_text}

Options (each with checkmark count shown in brackets):
{options_text}

Please determine which option is most likely selected as the correct answer. 
If multiple options have the same check count, choose the first option with that count.
Return only the option text.
Answer:"""

            answer = get_llm_response(prompt, client)
        else:
            answer = selected_options[0] if selected_options else "No answer detected"

        qa_pairs.append((question_text, answer))
    
    return qa_pairs

def main():
    input_docx = r"C:\Users\i7 11th\Desktop\checks.docx"  # Change this to your actual file path
    client = OpenAI(api_key="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")  # Replace with your OpenAI API key

    results = process_mcq_questions(input_docx, client)

    # Save the results
    with open("output.txt", "w", encoding="utf-8") as f:
        for q, a in results:
            f.write(f"Q: {q}\nA: {a}\n\n")

    print("✅ Results saved to output.txt")

if __name__ == "__main__":
    main()
