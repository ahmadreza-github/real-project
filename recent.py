import json
from groq import Groq # Don't forget to install groq -> pip install groq

def query_and_save():
    # Initialize Groq client
    client = Groq(api_key="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") # The api key is for llama-4, make sure you check results
    user_question = input("What is your question\n>>> ")

    user_prompt = f"""You are an LLM tutor for a 5th‑grade student. Use the following instructions to answer their questions:

                Content Of Instructions: "سلام دانشمند کوچک کلاس پنجمی! بیا با هم این آزمایش را که از درس "بکارید و بخورید" کتاب علوم است، قدم به قدم و با جزئیات کامل دنبال کنیم تا دقیقاً بفهمیم یک گیاه کوچولو مثل آفتابگردان برای رشد کردن به چه چیزهایی نیاز دارد...
                (متن کامل راهنما طبق نمونه قبلی که نوشتی اینجا قرار دارد)
                ...

                When the student asks a question, you must:
                1. Answer only based on the text inside those quotes.
                2. Reply entirely in Persian, no English.
                3. Use a friendly, age‑appropriate tone with examples or lists.
                4. If the question is outside the quoted instructions, say:  
                “متأسفم، در این مورد کمکی نمی‌توانم بکنم”  
                and remind them you can only help with این آزمایش کاشت آفتابگردان.
                here is the asked question: {user_question}
                """
    completion = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role": "user", "content": user_prompt}],
        temperature=1,
        max_tokens=1024,
        top_p=1,
        stream=True,
    )
    full_response = "" # full response will be saved here
    for chunk in completion:
        content = chunk.choices[0].delta.content or ""
        full_response += content

    # Save to 'new.json' as our saved data
    response_data = {
        "model": "meta-llama/llama-4-scout-17b-16e-instruct",
        "prompt": user_prompt,
        "reply": full_response
    }

    with open("new.json", "w", encoding="utf-8") as f_json:
        json.dump(response_data, f_json, ensure_ascii=False, indent=2)

    # Save only reply to a plain text file
    with open("response.txt", "w", encoding="utf-8") as f_txt:
        f_txt.write(full_response)

    print("\n>> Response saved to 'new.json' and 'response.txt' <<")

if __name__ == "__main__":
    query_and_save()


















