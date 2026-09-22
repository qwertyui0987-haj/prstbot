import json
import os
import io
import re
import asyncio
import urllib.request
import urllib.parse
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

def _get_gemini_response_sync(prompt: str) -> str:
    """Запрос через новый официальный SDK google-genai"""
    if not GEMINI_API_KEY:
        raise Exception("GEMINI_API_KEY не найден в переменных окружения!")

    client = genai.Client(api_key=GEMINI_API_KEY)
    
    # Актуальные рабочие модели
    candidate_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash"
    ]

    for model_name in candidate_models:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json"
                )
            )
            if response and response.text:
                return response.text
        except Exception as err:
            print(f"[GENERATOR LOG] Ошибка модели {model_name}: {err}")
            continue

    raise Exception("Ни одна из моделей Gemini не ответила.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Ushbu ssenariy bo'yicha EXACTLY 6 ta slayd yaratib ber.
        Javobingiz faqat quyidagi standart JSON massivi bo'lsin:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
            "image_keyword": "business"
          }}
        ]
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha EXACTLY 6 ta slaydli prezentatsiya yarat.
        Javobingiz faqat quyidagi standart JSON massivi bo'lsin:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
            "image_keyword": "technology"
          }}
        ]
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
    except Exception as e:
        print(f"[GENERATOR ERROR] Ошибка API: {e}")
        raw_response = ""

    # Парсинг JSON
    clean_text = raw_response.strip()
    clean_text = re.sub(r'^```(?:json)?\s*', '', clean_text, flags=re.IGNORECASE)
    clean_text = re.sub(r'\s*```$', '', clean_text, flags=re.IGNORECASE).strip()

    try:
        data = json.loads(clean_text)
        if isinstance(data, list) and len(data) > 0:
            return data
    except Exception:
        pass

    try:
        start_idx = clean_text.find('[')
        end_idx = clean_text.rfind(']')
        if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
            data = json.loads(clean_text[start_idx:end_idx + 1])
            if isinstance(data, list):
                return data
    except Exception:
        pass

    # Резервный вариант на случай сбоя API
    return [
        {
            "title": topic,
            "content": [
                f"{topic} - Kirish va umumiy tushunchalar",
                "Mavzuning asosiy yo'nalishlari va tahlili",
                "Xulosalar va amaliy ahamiyati"
            ],
            "image_keyword": "presentation"
        }
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
