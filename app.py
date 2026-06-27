from fastapi.responses import FileResponse
import os
from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request
from fastapi.middleware.cors import CORSMiddleware
import pdfplumber
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx
import json
import time  # 🧠 تم إضافة المكتبة هنا لحل مشكلة انهيار السيرفر الداخلي
import io
from google import genai
from google.genai import types
import sqlite3

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

def init_db():
    conn = sqlite3.connect("exam_platform.db")
    cursor = conn.cursor()
    # إنشاء جدول حفظ السجلات والنتائج إذا لم يكن موجوداً مسبقاً
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS exam_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            difficulty TEXT,
            language TEXT,
            score INTEGER,
            total_questions INTEGER,
            percentage REAL,
            timestamp TEXT
        )
    """)
    conn.commit()
    conn.close()

# استدعاء الدالة فوراً لتجهيز قاعدة البيانات في السيرفر
init_db()

MAX_REQUESTS = 3
TIME_WINDOW = 60

def extract_text_from_pdf(file_bytes):
    full_text = ""
    # القراءة من الذاكرة مباشرة لضمان عدم حدوث مشاكل في السيرفر
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
    return full_text

def extract_text_from_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    full_text = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(full_text)

def generate_questions(text_content, difficulty, num_mcq, num_tf, num_essay, language):
    lang_instruction = "باللغة العربية" if language == "ar" else "strictly in English language"
    
    prompt = f"""
    أنت بروفيسور وخبير تعليمي متخصص. بناءً على النص أدناه، قم بتوليد اختبار أكاديمي بدقة متناهية.
    
    مواصفات الاختبار المطلوبة:
    - لغة صياغة الأسئلة والخيارات بالكامل: يجب أن تكون {lang_instruction}.
    - مستوى الصعوبة: {difficulty}.
    - عدد أسئلة اختيار من متعدد: {num_mcq} أسئلة.
    - عدد أسئلة صح أو خطأ: {num_tf} أسئلة.
    - عدد الأسئلة المقالية: {num_essay} أسئلة.
    
     must return a valid JSON object strictly matching this schema:
    {{
      "multiple_choice": [
        {{"question": "نص السؤال؟", "options": ["1", "2", "3", "4"], "answer": "الخيار المطابق"}}
      ],
      "true_false": [
        {{"question": "نص السؤال؟", "answer": true}}
      ],
      "essay": [
        {{"question": "نص السؤال؟"}}
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
                detail="! لقد تجاوزت الحد المسموح للطلبات. انتظر دقيقة ثم حاول مجدداً"
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
    
    file_extension = file.filename.split(".")[-1].lower()
    
    try:
        if file_extension == "pdf":
            text = extract_text_from_pdf(file_contents)
        else:
            text = extract_text_from_docx(file_contents)
            
        if not text.strip():
            raise HTTPException(status_code=400, detail="الملف المرفوع فارغ أو لا يحتوي على نصوص قابلة للقراءة.")
            
        exam_json = generate_questions(text, difficulty, num_mcq, num_tf, num_essay, language)
        return exam_json
        
    except Exception as e:
        print("[❌] خطأ داخلي:", str(e))
        raise HTTPException(status_code=500, detail=f"حدث خطأ في النظام الداخلي: {str(e)}")

@app.post("/save-exam-result")
async def save_exam_result(data: dict):
    """🧠 مسار استقبال النتيجة النهائية من الواجهة وحفظها في قاعدة البيانات"""
    try:
        conn = sqlite3.connect("exam_platform.db")
        cursor = conn.cursor()
        
        # حقن البيانات بأمان باستخدام بروتوكول منسق لمنع هجمات SQL Injection
        cursor.execute("""
            INSERT INTO exam_logs (filename, difficulty, language, score, total_questions, percentage, timestamp)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            data.get("filename", "مستند غير مسمى"),
            data.get("difficulty", "غير محدد"),
            data.get("language", "ar"),
            data.get("score", 0),
            data.get("total_questions", 0),
            data.get("percentage", 0.0),
            time.strftime("%Y-%m-%d %H:%M:%S") # توثيق الوقت والتاريخ الحالي للطلب
        ))
        
        conn.commit()
        conn.close()
        return {"status": "success", "message": "تم توثيق وحفظ النتيجة في السجل الأكاديمي بأمان."}
    except Exception as e:
        print("[❌] خطأ في قاعدة البيانات:", str(e))
        raise HTTPException(status_code=500, detail=f"فشل حفظ النتيجة في قاعدة البيانات: {str(e)}")

@app.get("/get-exams-history")
async def get_exams_history():
    """📊 مسار جلب سجل آخر 5 اختبارات تم توليدها لعرضها في لوحة التحكم"""
    try:
        conn = sqlite3.connect("exam_platform.db")
        cursor = conn.cursor()
        
        # جلب آخر 5 اختبارات مرتبة من الأحدث إلى الأقدم
        cursor.execute("""
            SELECT filename, difficulty, language, score, total_questions, percentage, timestamp 
            FROM exam_logs 
            ORDER BY id DESC 
            LIMIT 5
        """)
        rows = cursor.fetchall()
        conn.close()
        
        # تحويل السطور المجلوبة إلى قائمة كائنات JSON مفهومة للمتصفح
        history_list = []
        for row in rows:
            history_list.append({
                "filename": row[0],
                "difficulty": row[1],
                "language": row[2],
                "score": row[3],
                "total_questions": row[4],
                "percentage": row[5],
                "timestamp": row[6]
            })
        return history_list
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"فشل جلب السجلات: {str(e)}")


@app.get("/")
async def serve_frontend():
    return FileResponse("index.html")

@app.post("/export-exam")
async def export_exam_to_docx(exam_data: dict):
    doc = Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # الترويسة الأكاديمية
    title = doc.add_paragraph()
    title_run = title.add_run("جامعة: ............................\nالكلية: ............................\nالقسم: ............................")
    title_run.font.name = 'Arial'
    title_run.font.size = Pt(12)
    title_run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    header_exam = doc.add_paragraph()
    header_run = header_exam.add_run("\nإمتحان المادة التفاعلي الذكي\nالزمن: ساعتان\n")
    header_run.font.name = 'Arial'
    header_run.font.size = Pt(14)
    header_run.bold = True
    header_exam.alignment = WD_ALIGN_PARAGRAPH.CENTER

    student_info = doc.add_paragraph()
    student_run = student_info.add_run("اسم الطالب: ............................................................  الرقم الأكاديمي: .............................")
    student_run.font.name = 'Arial'
    student_run.font.size = Pt(11)
    student_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("------------------------------------------------------------------------------------------------------------------------")

    global_index = 1

    # 1. طباعة أسئلة الاختيار من متعدد
    if "multiple_choice" in exam_data and exam_data["multiple_choice"]:
        h = doc.add_paragraph()
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        h.add_run("القسم الأول: أسئلة الاختيار من متعدد").bold = True
        
        for q in exam_data["multiple_choice"]:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"س {global_index}: {q.get('question')}")
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            
            # طباعة الخيارات
            options = q.get("options", [])
            for idx, opt in enumerate(options, 1):
                opt_p = doc.add_paragraph()
                opt_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                opt_p.add_run(f"   [{idx}] {opt}").font.name = 'Arial'
            global_index += 1
        doc.add_paragraph("\n")

    # 2. طباعة أسئلة صح أم خطأ
    if "true_false" in exam_data and exam_data["true_false"]:
        h = doc.add_paragraph()
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        h.add_run("القسم الثاني: أسئلة صح أم خطأ (ضع علامة صح أو خطأ أمام العبارات التالية)").bold = True
        
        for q in exam_data["true_false"]:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"س {global_index}: {q.get('question')}  (   )")
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            global_index += 1
        doc.add_paragraph("\n")

    # 3. طباعة الأسئلة المقالية
    if "essay" in exam_data and exam_data["essay"]:
        h = doc.add_paragraph()
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        h.add_run("القسم الثالث: الأسئلة المقالية (أجب عن الأسئلة التالية بالتفصيل)").bold = True
        
        for q in exam_data["essay"]:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(f"س {global_index}: {q.get('question')}")
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            doc.add_paragraph("\nالإجابة:\n........................................................................................................................\n........................................................................................................................")
            global_index += 1

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=Generated_Exam.docx"}
    )
