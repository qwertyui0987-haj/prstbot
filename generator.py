import json
import os
import io
import asyncio
import urllib.request
import urllib.parse
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

def _get_gemini_response_sync(prompt: str) -> str:
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Berilgan ssenariy bo'yicha 6 ta mukammal va mazmunli slayd yarat.
        Har bir slayd uchun inglizcha rasm kalit so'zini (image_keyword) ko'rsat.
        
        Javobni FAQAT JSON formatida qaytar:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["To'liq ma'lumotli punkt 1", "To'liq ma'lumotli punkt 2", "To'liq ma'lumotli punkt 3"],
            "image_keyword": "nature"
          }}
        ]
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzuda 6 ta batafsil va professional slayd tayyorla.
        Matnlar o'zbek tilida, tushunarli va boy mazmunga ega bo'lsin.
        Har bir slayd uchun inglizcha mos rasm kalit so'zini (image_keyword) ham ber.
        
        Javobni FAQAT JSON formatida qaytar:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Mavzuga oid batafsil fikr 1", "Mavzuga oid batafsil fikr 2", "Mavzuga oid batafsil fikr 3"],
            "image_keyword": "environment"
          }}
        ]
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = raw_response.strip()
        
        if "```json" in clean_text:
            clean_text = clean_text.split("```json")[1].split("```")[0].strip()
        elif "```" in clean_text:
            clean_text = clean_text.split("```")[1].split("```")[0].strip()
            
        return json.loads(clean_text)
    except Exception as e:
        print(f"Gemini API xatosi: {e}")
        return [
            {
                "title": topic,
                "content": ["Kirish qismi va mavzu bo'yicha umumiy ma'lumotlar.", "Asosiy tushunchalar hamda ularning tahlili."],
                "image_keyword": "presentation"
            }
        ]

def _fetch_image_sync(keyword: str):
    """Mavzuga mos rasmlarni yuklab olish"""
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        url = f"https://source.unsplash.com/800x600/?{encoded_keyword}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as resp:
            return io.BytesIO(resp.read())
    except Exception:
        return None

def _build_pptx_sync(slides_data: list, output_filename: str) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    PRIMARY_COLOR = RGBColor(20, 35, 60)
    TEXT_COLOR = RGBColor(40, 40, 40)

    for i, slide_info in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # 1-Slayd: Titul (AI haqida hech qanday yozuvsiz)
        if i == 0:
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.8), Inches(11.333), Inches(2.0))
            tf = title_box.text_frame
            tf.word_wrap = True
            
            p = tf.paragraphs[0]
            p.text = slide_info.get("title", "Prezentatsiya")
            p.font.size = Pt(48)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            p.alignment = PP_ALIGN.CENTER
            continue

        # Slayd sarlavhasi (Katta va aniq)
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", "")
        p_title.font.size = Pt(34)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Slayd matni (Shrift o'lchami 22pt)
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(22)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(16)

        # Rasm (O'ng tomonda)
        keyword = slide_info.get("image_keyword", "topic")
        img_stream = _fetch_image_sync(keyword)
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(8.0), Inches(1.8), width=Inches(4.5))
            except Exception:
                pass

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
