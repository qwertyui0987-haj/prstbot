import json
import os
import io
import asyncio
import urllib.request
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

def _get_gemini_response_sync(prompt: str) -> str:
    """Gemini API so'rovini bajarish"""
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Berilgan ssenariyni 6-8 ta chiroyli, mantiqiy va xatosiz o'zbek tilidagi slaydlarga bo'lib bering.
        Har bir slayd uchun:
        - title: Slayd sarlavhasi
        - content: 3-4 ta asosiy va batafsil fikrlar
        - image_keyword: Slaydga mos inglizcha bitta rasm kalit so'zi (masalan: nature, ecology, technology)

        Javobni FAQAT quyidagi JSON formatida qaytaring, ortiqcha markdown va matnlarsiz:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Batafsil fikr 1", "Batafsil fikr 2", "Batafsil fikr 3"],
            "image_keyword": "nature"
          }}
        ]
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha professional va to'liq 6-8 ta slayd tayyorlang.
        Barcha matnlar o'zbek tilida gramatik xatosiz, ilmiy va tushunarli bo'lsin.
        Har bir slayd uchun:
        - title: Slayd sarlavhasi
        - content: 3-5 ta batafsil ma'lumot beruvchi punktlar
        - image_keyword: Slaydga mos inglizcha bitta rasm kalit so'zi (masalan: environment, save Earth, green energy)

        Javobni FAQAT quyidagi JSON formatida qaytaring, ortiqcha markdown va matnlarsiz:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Batafsil punkt 1", "Batafsil punkt 2", "Batafsil punkt 3"],
            "image_keyword": "nature"
          }}
        ]
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = raw_response.strip().replace("```json", "").replace("```", "")
        slides_data = json.loads(clean_text)
        return slides_data
    except Exception as e:
        print(f"Gemini API / JSON Xatosi: {e}")
        # Zaxira uchun kengaytirilgan slaydlar
        return [
            {
                "title": f"{topic}",
                "content": ["Kirish va umumiy tushuncha", "Mavzuning dolzarbligi va ahamiyati"],
                "image_keyword": "presentation"
            },
            {
                "title": "Asosiy Maqsad va Vazifalar",
                "content": ["Muammoni o'rganish va tahlil qilish", "Yechimlar hamda amaliy takliflar berish"],
                "image_keyword": "target"
            },
            {
                "title": "Xulosa va Tavsiyalar",
                "content": ["Erishilgan asosiy natijalar", "Kelajakdagi istiqbolli qadamlar"],
                "image_keyword": "success"
            }
        ]

def _fetch_image(keyword: str):
    """Unsplash orqali mavzuga mos bepul rasm yuklab olish"""
    try:
        url = f"https://source.unsplash.com/800x600/?{keyword}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as resp:
            return io.BytesIO(resp.read())
    except Exception:
        return None

def _build_pptx_sync(slides_data: list, output_filename: str) -> str:
    """Chiroyli va rasmli PPTX faylini shakllantirish"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    PRIMARY_COLOR = RGBColor(15, 32, 67)    # To'q ko'k
    ACCENT_COLOR = RGBColor(0, 122, 255)    # Moviy
    TEXT_COLOR = RGBColor(50, 50, 50)

    for i, slide_info in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # 1-Slayd: Muqova (Title Slide)
        if i == 0:
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.0))
            tf = title_box.text_frame
            tf.word_wrap = True
            
            p = tf.paragraphs[0]
            p.text = slide_info.get("title", "Prezentatsiya")
            p.font.size = Pt(44)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            p.alignment = PP_ALIGN.CENTER
            
            p2 = tf.add_paragraph()
            p2.text = "\nTayyorladi: Sun'iy Intellekt (Gemini AI)"
            p2.font.size = Pt(20)
            p2.font.color.rgb = ACCENT_COLOR
            p2.alignment = PP_ALIGN.CENTER
            continue

        # Boshqa slaydlar: Sarlavha + Matn + Rasm
        # Sarlavha
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"Slayd {i+1}")
        p_title.font.size = Pt(30)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Matn bloki (Chap tomonda)
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(7.0), Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(18)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(12)

        # Rasm qo'shish (O'ng tomonda)
        img_keyword = slide_info.get("image_keyword", "nature")
        img_stream = _fetch_image(img_keyword)
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(8.2), Inches(1.8), width=Inches(4.3))
            except Exception as e:
                print(f"Rasm joylashda xato: {e}")

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
