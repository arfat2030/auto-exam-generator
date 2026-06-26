from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
import pdfplumber
import docx
import json
from google import genai
from google.genai import types

app = FastAPI(title="Auto Exam Generator API - V3")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY)

def extract_text_from_pdf(file_object):
    full_text = ""
    with pdfplumber.open(file_object) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
    return full_text

def extract_text_from_docx(file_object):
    doc = docx.Document(file_object)
    full_text = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(full_text)

# أضفنا معامل لغة الاختبار هنا
def generate_questions(text_content, difficulty, num_mcq, num_tf, num_essay, language):
    # نحدد للذكاء الاصطناعي اللغة المطلوبة صراحة
    lang_instruction = "باللغة العربية" if language == "ar" else "strictly in English language"
    
    prompt = f"""
    أنت بروفيسور وخبير تعليمي متخصص. بناءً على النص أدناه، قم بتوليد اختبار أكاديمي بدقة متناهية.
    
    مواصفات الاختبار المطلوبة:
    - لغة صياغة الأسئلة والخيارات بالكامل: يجب أن تكون {lang_instruction}.
    - مستوى الصعوبة: {difficulty}.
    - عدد أسئلة اختيار من متعدد: {num_mcq} أسئلة.
    - عدد أسئلة صح أو خطأ: {num_tf} أسئلة.
    - عدد الأسئلة المقالية: {num_essay} أسئلة.
    
    يجب أن تكون الإجابة بصيغة JSON حصراً وبنفس البنية تماماً دون أي نص ترحيبي خارجها:
    {{
      "multiple_choice": [
        {{"question": "نص السؤال باللغة المختارة؟", "options": ["1", "2", "3", "4"], "answer": "الخيار المطابق تماماً"}}
      ],
      "true_false": [
        {{"question": "نص السؤال باللغة المختارة؟", "answer": true}}
      ],
      "essay": [
        {{"question": "نص السؤال باللغة المختارة؟"}}
      ]
    }}
    
    النص المراد بناء الاختبار منه:
    \"\"\"
    {text_content}
    \"\"\"
    """
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=prompt,
        config=types.GenerateContentConfig(response_mime_type="application/json"),
    )
    return json.loads(response.text)

@app.post("/generate-exam")
async def generate_exam_endpoint(
    file: UploadFile = File(...),
    difficulty: str = Form(...),
    num_mcq: int = Form(...),
    num_tf: int = Form(...),
    num_essay: int = Form(...),
    language: str = Form(...) # استقبال اللغة من الواجهة
):
    file_extension = file.filename.split(".")[-1].lower()
    if file_extension not in ["pdf", "docx"]:
        raise HTTPException(status_code=400, detail="عذراً، يجب رفع ملف PDF أو Word فقط.")
    
    try:
        file_content = file.file
        if file_extension == "pdf":
            text = extract_text_from_pdf(file_content)
        else:
            text = extract_text_from_docx(file_content)
            
        if not text.strip():
            raise HTTPException(status_code=400, detail="الملف المرفوع فارغ.")
            
        # تمرير اللغة المحددة للدالة
        exam_json = generate_questions(text, difficulty, num_mcq, num_tf, num_essay, language)
        return exam_json
        
    except Exception as e:
        print("[❌] خطأ داخلي:", str(e))
        raise HTTPException(status_code=500, detail=f"حدث خطأ في النظام الداخلي: {str(e)}")