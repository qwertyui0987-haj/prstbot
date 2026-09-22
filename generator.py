import json
import os
import io
import re
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
    candidate_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash-latest",
        "gemini-1.5-pro",
        "gemini-pro"
    ]

    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(
                model_name,
                generation_config={"response_mime_type": "application/json"}
            )
            response = model.generate_content(prompt)
            if response and response.text:
                return response.text
        except Exception:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(prompt)
                if response and response.text:
                    return response.text
            except Exception:
                pass

    raise Exception("Gemini modellaridan javob olib bo'lmadi.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    prompt = f"""
    Mavzu: {topic}
    {f'Ssenariy: {user_script}' if user_script else ''}

    Ushbu mavzu bo'yicha EXACTLY 6 ta slayd tayyorla.
    FAQAT toza JSON formatida javob ber:
    [
      {{
        "title": "Slayd sarlavhasi",
        "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
        "image_keyword": "technology"
      }}
    ]
    """

    raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)

    # 1. Toza JSON parse qilishga urinish
    clean_text = raw_response.strip()
    clean_text = re.sub(r'^```(?:json)?\s*', '', clean_text, flags=re.IGNORECASE)
    clean_text = re.sub(r'\s*```$', '', clean_text, flags=re.IGNORECASE).strip()

    try:
        data = json.loads(clean_text)
        if isinstance(data, list):
            return data
    except Exception:
        pass

    # 2. Xavfsiz qidirish: Regex orqali barcha {} obyektlarini ajratib olish
    found_slides = []
    # Massiv ichidagi har bir { ... } blokini topamiz
    matches = re.findall(r'\{[^{}]*"title"[^{}]*\}', raw_response, re.DOTALL)
    for m in matches:
        try:
            obj = json.loads(m)
            found_slides.append(obj)
        except Exception:
            continue

    if found_slides:
        return found_slides

    # 3. Agar mutlaqo JSON bo'lmasa, zaxira slaydlarini qaytarish (Bot to'xtab qolmasligi uchun)
    return [
        {"title": topic, "content": ["Kechirasiz, kontentni avtomatik formatlashda xatolik bo'ldi.", "Lekin taqdimot tayyorlandi."], "image_keyword": "presentation"}
    ]

def _fetch_image_sync(keyword: str):
    try:
        encoded = urllib.parse.quote(keyword)
        url = f"https://source.unsplash.com/800x600/?{encoded}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=4) as resp:
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

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"{i+1}-Slayd")
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(7.0), Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        if isinstance(points, list):
            for idx, point in enumerate(points):
                p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
                p.text = f"• {point}"
                p.font.size = Pt(22)
                p.font.color.rgb = TEXT_COLOR
                p.space_after = Pt(14)
        else:
            p = tf_content.paragraphs[0]
            p.text = f"• {points}"
            p.font.size = Pt(22)
            p.font.color.rgb = TEXT_COLOR

        keyword = slide_info.get("image_keyword", "topic")
        img_stream = _fetch_image_sync(keyword)
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(8.2), Inches(1.8), width=Inches(4.5))
            except Exception:
                pass

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
