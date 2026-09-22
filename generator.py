import json
import os
import asyncio
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

def _get_gemini_response_sync(prompt: str) -> str:
    """Gemini AI so'rovini sinxron bajaruvchi yordamchi funksiya"""
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Berilgan ssenariyni 5-8 ta chiroyli va xatosiz o'zbek tilidagi slaydlarga bo'lib bering.
        Har bir slayd uchun sarlavha (title) va 3-4 ta asosiy fikr (content) tayyorlang.
        
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
        Mavzu: {topic}

        Vazifa: 6-8 ta slayd tayyorlang. Barcha matnlar o'zbek tilida gramatik xatosiz bo'lsin.
        Har bir slayd uchun qisqa sarlavha (title) va 3-5 ta punktlar (content) yozing.

        Javobni FAQAT quyidagi JSON formatida qaytaring, ortiqcha matn yoki markdown belgilarisiz:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Fikr 1", "Fikr 2", "Fikr 3"]
          }}
        ]
        """

    try:
        # Gemini API so'rovini alohida oqimda bajarish (loop qotib qolmasligi uchun)
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = raw_response.strip().replace("```json", "").replace("```", "")
        slides_data = json.loads(clean_text)
        return slides_data
    except Exception as e:
        print(f"Gemini/JSON Xatosi: {e}")
        return [
            {"title": topic, "content": ["Prezentatsiya avtomatik yaratildi."]},
            {"title": "Xulosa", "content": ["E'tiboringiz uchun rahmat!"]}
        ]

def _build_pptx_sync(slides_data: list, output_filename: str) -> str:
    """PPTX faylini yaratish"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    PRIMARY_COLOR = RGBColor(24, 43, 73)
    TEXT_COLOR = RGBColor(40, 40, 40)

    for i, slide_info in enumerate(slides_data):
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.2))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"Slayd {i+1}")
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.7), Inches(4.8))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"•  {point}"
            p.font.size = Pt(20)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(14)

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    # Fayl yozishni ham alohida thread'ga olamiz
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
