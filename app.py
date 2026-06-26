from fastapi.responses import FileResponse
import os
from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi import Request
import pdfplumber
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx
import json
from google import genai
from google.genai import types

app = FastAPI(title="Auto Exam Generator API - V3")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False, 
    allow_methods=["*"],
    allow_headers=["*"],
)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY)

REQUESTS_TRACKER = {}

# الإعدادات الأمنية: 3 طلبات كحد أقصى خلال 60 ثانية
MAX_REQUESTS = 3
TIME_WINDOW = 60

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
    request: Request,
    file: UploadFile = File(...),
    difficulty: str = Form(...),
    num_mcq: int = Form(...),
    num_tf: int = Form(...),
    num_essay: int = Form(...),
    language: str = Form(...) 
):
    client_ip = request.client.host 
    current_time = time.time() 
    
    
    if client_ip in REQUESTS_TRACKER:
        
        REQUESTS_TRACKER[client_ip] = [
            t for t in REQUESTS_TRACKER[client_ip] if current_time - t < TIME_WINDOW
        ]

        if len(REQUESTS_TRACKER[client_ip]) >= MAX_REQUESTS:
            raise HTTPException(
                status_code=429, 
                detail=" !لقد تجاوزت الحد المسموح للطلبات. انتظر دقيقة ثم حاول مجدداً"
            )

        REQUESTS_TRACKER[client_ip].append(current_time)
    else:
        REQUESTS_TRACKER[client_ip] = [current_time]
    ALLOWED_EXTENSIONS = ('.pdf', '.docx')
    if not file.filename.lower().endswith(ALLOWED_EXTENSIONS):
        raise HTTPException(
            status_code=400, 
            detail="عذراً، النظام يقبل ملفات PDF وملفات Word (.docx) فقط."
        )

    MAX_FILE_SIZE = 10 * 1024 * 1024  
    
    file_contents = await file.read()
    file_size = len(file_contents)
    
    if file_size > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=400, 
            detail="الملف ضخم جداً! لأسباب أمنية وحفاظاً على موارد السيرفر، الحد الأقصى المسموح به هو 10 ميجابايت."
        )
        
    await file.seek(0)
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
            
        
        exam_json = generate_questions(text, difficulty, num_mcq, num_tf, num_essay, language)
        return exam_json
        
    except Exception as e:
        print("[❌] خطأ داخلي:", str(e))
        raise HTTPException(status_code=500, detail=f"حدث خطأ في النظام الداخلي: {str(e)}")
        
@app.get("/")
async def serve_frontend():
    return FileResponse("index.html")

@app.post("/export-exam")
async def export_exam_to_docx(exam_data: dict):
    #  إنشاء مستند Word جديد في الذاكرة
    doc = Document()
    
    # 📝 ضبط اتجاه الصفحة ليدعم اللغة العربية (من اليمين إلى اليسار)
    # ملاحظة: ملفات الـ Word تحتاج تنسيقاً أساسياً للفقرات
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    #  1. بناء ترويسة الاختبار الأكاديمية (Header)
    title = doc.add_paragraph()
    title_run = title.add_run("جامعة: ............................\nالكلية: ............................\nالقسم: ............................")
    title_run.font.name = 'Arial'
    title_run.font.size = Pt(12)
    title_run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 📄 عنوان المادة والامتحان في المنتصف
    header_exam = doc.add_paragraph()
    header_run = header_exam.add_run("\nإمتحان المادة التفاعلي النهائي\nالزمن: ساعتان\n")
    header_run.font.name = 'Arial'
    header_run.font.size = Pt(14)
    header_run.bold = True
    header_exam.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #  سطر بيانات الطالب
    student_info = doc.add_paragraph()
    student_run = student_info.add_run("اسم الطالب: ............................................................  الرقم الأكاديمي: .............................")
    student_run.font.name = 'Arial'
    student_run.font.size = Pt(11)
    student_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("------------------------------------------------------------------------------------------------------------------------")

    #  2. قراءة الأسئلة من الـ JSON وكتابتها داخل الملف
    questions = exam_data.get("questions", [])
    
    for index, q in enumerate(questions, 1):
        q_type = q.get("type", "")
        q_text = q.get("question", "")
        
        # كتابة نص السؤال
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(f"س {index}: {q_text}")
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        
        if q_type == "mcq" and "options" in q:
            for opt_key, opt_val in q["options"].items():
                opt_p = doc.add_paragraph()
                opt_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                opt_run = opt_p.add_run(f"   {opt_key}) {opt_val}")
                opt_run.font.name = 'Arial'
                opt_run.font.size = Pt(11)
                
        elif q_type == "essay":
            doc.add_paragraph("\nالإجابة:\n........................................................................................................................\n........................................................................................................................") 
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    #  إرسال الملف فوراً للمتصفح ليتم تحميله باسم رسمي تلقائي
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=Generated_Exam.docx"}
    )
