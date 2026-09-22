import json
import os
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from dotenv import load_dotenv

load_dotenv()

# Gemini AI-ni sozlash
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    """
    Gemini AI orqali slaydlar strukturasini JSON formatida oladi.
    """
    model = genai.GenerativeModel("gemini-1.5-flash")
    
    if user_script:
        prompt = f"""
        Foydalanuvchi quyidagi mavzu va ssenariyni taqdim etdi:
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifangiz: Berilgan ssenariyni 5-8 ta chiroyli, mantiqiy va xatosiz o'zbek tilidagi slaydlarga bo'lib bering.
        Har bir slayd uchun sarlavha (title) va 3-4 ta asosiy fikr (bullet_points) tayyorlang.
        
        Javobni FAQAT quyidagi JSON formatida qaytaring, ortiqcha matn yoki markdown belgilarisiz:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Fikr 1", "Fikr 2", "Fikr 3"]
          }}
        ]
        """
    else:
        prompt = f"""
        Siz professional prezentatsiya mutaxassisiz. Quyidagi mavzu bo'yicha to'liq prezentatsiya strukturasi va matnlarini tuzing:
        Mavzu: {topic}

        Vazifangiz:
        1. 6-10 ta slayd tayyorlang (Kirish, Asosiy qism, Xulosa/Xulosa tavsiyalari).
        2. Barcha matnlar o'zbek tilida gramatik xatosiz va akademik/professional uslubda bo'lsin.
        3. Har bir slayd uchun qisqa sarlavha (title) va 3-5 ta tushunarli punktlar (bullet_points) yozing.

        Javobni FAQAT quyidagi JSON formatida qaytaring, ortiqcha matn yoki markdown belgilarisiz:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Fikr 1", "Fikr 2", "Fikr 3"]
          }}
        ]
        """

    response = await model.generate_content_async(prompt)
    clean_text = response.text.strip().replace("```json", "").replace("```", "")
    
    try:
        slides_data = json.loads(clean_text)
        return slides_data
    except Exception as e:
        print(f"JSON parsing xatosi: {e}")
        # Xatolik bo'lsa zaxira strukturasi
        return [
            {"title": topic, "content": ["Prezentatsiya avtomatik yaratildi."]},
            {"title": "Xulosa", "content": ["E'tiboringiz uchun rahmat!"]}
        ]


def create_pptx_file(slides_data: list, output_filename: str) -> str:
    """
    JSON ma'lumotlar asosida zamonaviy va chiroyli PPTX faylini yaratadi.
    """
    prs = Presentation()
    
    # Slayd o'lchamini 16:9 standartiga o'tkazish
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Rangi: To'q ko'k va zamonaviy uslub
    PRIMARY_COLOR = RGBColor(24, 43, 73)
    TEXT_COLOR = RGBColor(40, 40, 40)

    for i, slide_info in enumerate(slides_data):
        # Bo'sh slayd shabloni
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)

        # 1. Title (Sarlavha) joylashtirish
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.2))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"Slayd {i+1}")
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR
        p_title.font.name = "Calibri"

        # 2. Content (Matnlar) joylashtirish
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.7), Inches(4.8))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"•  {point}"
            p.font.size = Pt(20)
            p.font.color.rgb = TEXT_COLOR
            p.font.name = "Calibri"
            p.space_after = Pt(14)

    prs.save(output_filename)
    return output_filename