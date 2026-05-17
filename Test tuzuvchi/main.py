import os
import io
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from pypdf import PdfReader
import docx2txt  # Word fayllarini o'qish uchun
from google import genai
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = FastAPI(title="Smart Test Generator")

# 1. Gemini API Sozlamasi
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Agar kalit umuman topilmasa, dastur ishga tushayotgandayoq terminalda ogohlantiradi
if not GEMINI_API_KEY:
    print("DIQQAT: GEMINI_API_KEY muhit o'zgaruvchisi topilmadi! Tizim test tuza olmaydi.")

client = genai.Client(api_key=GEMINI_API_KEY)

# 2. Yangilangan Frontend interfeysi (PDF va Word uchun moslashtirilgan)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI Test Generator | CodeCraft</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <style>
        body {
            background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);
        }
        .glass-card {
            background: rgba(255, 255, 255, 0.03);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.08);
        }
    </style>
</head>
<body class="text-gray-100 min-h-screen flex items-center justify-center p-4">

    <div class="glass-card w-full max-w-2xl rounded-3xl p-8 shadow-2xl">
        <div class="text-center mb-8">
            <h1 class="text-3xl font-extrabold bg-gradient-to-r from-blue-400 to-indigo-400 bg-clip-text text-transparent">
                Smart Test Generator
            </h1>
            <p class="text-gray-400 text-sm mt-2">PDF yoki Word ma'ruzadan bir lahzada akademik testlar tayyorlang</p>
        </div>

        <form action="/generate" method="POST" enctype="multipart/form-data" class="space-y-6" onsubmit="showLoader()">
            <div class="border-2 border-dashed border-gray-600 hover:border-indigo-500 rounded-2xl p-6 text-center cursor-pointer transition">
                <input type="file" name="file" accept=".pdf,.docx,.doc" class="hidden" id="fileInput" required onchange="updateFileName()">
                <label for="fileInput" class="cursor-pointer space-y-2">
                    <svg class="mx-auto h-12 w-12 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                    </svg>
                    <p id="uploadText" class="text-base font-medium">PDF yoki Word (DOCX) faylni yuklang</p>
                    <p class="text-xs text-gray-500">Maksimal hajm: 20MB</p>
                </label>
            </div>

            <div>
                <label class="block text-sm font-semibold text-gray-300 mb-3">Savollar sonini tanlang:</label>
                <div class="grid grid-cols-4 gap-3">
                    <label class="cursor-pointer">
                        <input type="radio" name="question_count" value="10" class="peer hidden" checked>
                        <div class="glass-card text-center py-3 rounded-xl font-bold peer-checked:bg-indigo-600 peer-checked:text-white border border-gray-700 transition">10 ta</div>
                    </label>
                    <label class="cursor-pointer">
                        <input type="radio" name="question_count" value="15" class="peer hidden">
                        <div class="glass-card text-center py-3 rounded-xl font-bold peer-checked:bg-indigo-600 peer-checked:text-white border border-gray-700 transition">15 ta</div>
                    </label>
                    <label class="cursor-pointer">
                        <input type="radio" name="question_count" value="30" class="peer hidden">
                        <div class="glass-card text-center py-3 rounded-xl font-bold peer-checked:bg-indigo-600 peer-checked:text-white border border-gray-700 transition">30 ta</div>
                    </label>
                    <label class="cursor-pointer">
                        <input type="radio" name="question_count" value="50" class="peer hidden">
                        <div class="glass-card text-center py-3 rounded-xl font-bold peer-checked:bg-indigo-600 peer-checked:text-white border border-gray-700 transition">50 ta</div>
                    </label>
                </div>
            </div>

            <button id="submitBtn" type="submit" class="w-full bg-gradient-to-r from-blue-500 to-indigo-600 hover:from-blue-600 hover:to-indigo-700 text-white font-bold py-4 px-6 rounded-xl shadow-lg transition transform hover:-translate-y-0.5 flex justify-center items-center gap-2">
                Testni Generatsiya Qilish & PDF Yuklash
            </button>
        </form>
    </div>

    <script>
        function updateFileName() {
            const input = document.getElementById('fileInput');
            const text = document.getElementById('uploadText');
            if(input.files.length > 0) {
                text.innerText = "Yuklangan fayl: " + input.files[0].name;
                text.classList.add("text-green-400");
            }
        }

        function showLoader() {
            const btn = document.getElementById('submitBtn');
            btn.disabled = true;
            btn.innerHTML = `<svg class="animate-spin h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg> AI test tuzmoqda, iltimos kuting...`;
            setTimeout(() => {
                btn.disabled = false;
                btn.innerHTML = "Testni Generatsiya Qilish & PDF Yuklash";
            }, 12000);
        }
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def main_page():
    return HTML_TEMPLATE


# 3. Test generatsiya qilish va fayllarni qayta ishlash logikasi
@app.post("/generate")
async def generate_tests(
    file: UploadFile = File(...), 
    question_count: int = Form(...)
):
    filename = file.filename.lower()
    extracted_text = ""

    try:
        # Fayl tarkibini o'qib olish
        file_bytes = await file.read()
        file_like_object = io.BytesIO(file_bytes)

        # a) Agar fayl PDF bo'lsa
        if filename.endswith('.pdf'):
            pdf_reader = PdfReader(file_like_object)
            for page in pdf_reader.pages:
                extracted_text += page.extract_text() + "\n"

        # b) Agar fayl Word (DOCX) bo'lsa
        elif filename.endswith('.docx'):
            extracted_text = docx2txt.process(file_like_object)

        # d) Noto'g'ri format kiritilsa xatolik qaytarish
        else:
            raise HTTPException(status_code=400, detail="Faqat .pdf yoki .docx formatidagi fayllarni yuklashingiz mumkin!")

        # Matn bo'shligini tekshirish
        if not extracted_text.strip():
            raise HTTPException(status_code=400, detail="Yuklangan fayldan matnni ajratib bo'lmadi. Fayl bo'sh emasligiga ishonch hosil qiling.")

        # Matn limitini sozlash (boshlang'ich tahlil uchun 6000 belgi)
        context_text = extracted_text[:6000]

        # c) Gemini API orqali test yaratish prompti
        prompt = f"""
        Siz oliy ta'lim muassasasi professorsiz. Berilgan matn mazmunidan kelib chiqib, aniq {question_count} ta test savoli tayyorlang.
        
        Qat'iy talablar:
        1. Savollar soni kam ham, ko'p ham bo'lmasin, aynan {question_count} ta bo'lsin.
        2. Har bir savolda 4 ta muqobil variant (A, B, C, D) bo'lsin va ulardan faqat bittasi to'g'ri bo'lsin.
        3. Savollar takrorlanmasin, qisqa, lo'nda va ilmiy tilda bo'lsin.
        4. Formatni aniq quyidagicha saqlang:
           1-savol matni...
           A) variant  B) variant  C) variant  D) variant
        5. Barcha testlar tugagach, eng pastda yangi qatordan "TO'G'RI JAVOBLAR KALITI" sarlavhasini oching va javoblarni "1-A, 2-C, 3-D..." shaklida ketma-ket yozing.
        
        Matn:
        {context_text}
        """

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        ai_result = response.text

        # e) Natijani PDF formatga eksport qilish (In-Memory)
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        story = []

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=15,
            alignment=1
        )
        body_style = ParagraphStyle(
            'BodyStyle',
            parent=styles['Normal'],
            fontSize=10,
            leading=14,
            spaceAfter=8
        )

        story.append(Paragraph(f"<b>FAN BO'YICHA TEST TOPSHIRIQLARI ({question_count} talik)</b>", title_style))
        story.append(Spacer(1, 10))

        lines = ai_result.split('\n')
        for line in lines:
            if line.strip():
                story.append(Paragraph(line, body_style))

        doc.build(story)
        pdf_buffer.seek(0)

        # f) Tayyor PDFni foydalanuvchiga taqdim etish
        return StreamingResponse(
            pdf_buffer, 
            media_type="application/pdf", 
            headers={"Content-Disposition": f"attachment; filename=Smart_Test_{question_count}.pdf"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Xatolik yuz berdi: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)