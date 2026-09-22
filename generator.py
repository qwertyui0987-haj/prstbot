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
else:
    print("CRITICAL ERROR: GEMINI_API_KEY muhit o'zgaruvchisi olinmadi!")

def _get_gemini_response_sync(prompt: str) -> str:
    """Gemini modelidan matn olish"""
    candidate_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash-latest",
        "gemini-1.5-pro",
        "gemini-pro"
    ]

    last_exception = None

    for model_name in candidate_models:
        try:
            print(f"[GENERATOR] Model sinab ko'rilmoqda: {model_name}")
            model = genai.GenerativeModel(
                model_name,
                generation_config={"response_mime_type": "application/json"}
            )
            response = model.generate_content(prompt)
            if response and response.text:
                print(f"[GENERATOR] Muvaffaqiyatli model: {model_name}")
                return response.text
        except Exception as err:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(prompt)
                if response and response.text:
                    return response.text
            except Exception:
                pass
            print(f"[GENERATOR] {model_name} xatosi: {err}")
            last_exception = err

    if last_exception:
        raise last_exception
    raise Exception("Birorta ham ishlaydigan Gemini modeli topilmadi.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    print(f"\n==========================================")
    print(f"[GENERATOR] Funksiya chaqirildi! Mavzu: '{topic}'")
    print(f"==========================================\n")

    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Berilgan ssenariy bo'yicha 6 ta slayd tayyorla. 
        Javobingiz FAQAT JSON formatidagi massiv bo'lsin:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
            "image_keyword": "nature"
          }}
        ]
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Ushbu mavzu bo'yicha 6 ta professional slayd tayyorla.
        Javobingiz FAQAT JSON formatidagi massiv bo'lsin:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Mavzuga oid batafsil fikr 1", "Mavzuga oid batafsil fikr 2", "Mavzuga oid batafsil fikr 3"],
            "image_keyword": "technology"
          }}
        ]
        """

    raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
    print(f"[GENERATOR] Gemini javob berdi! Uzunligi: {len(raw_response)}")

    # REGEX ORQALI HAR BIR SLAYD OBYEKTINI ALOHIDA SUZIB OLAMIZ (Extra Data xatosini 100% yo'qotadi)
    dict_matches = re.findall(r'\{[^{}]*"title"[^{}]*\}', raw_response, re.DOTALL)
    
    slides_list = []
    if dict_matches:
        for match in dict_matches:
            try:
                slide_obj = json.loads(match)
                slides_list.append(slide_obj)
            except Exception:
                continue

    if slides_list:
        return slides_list

    # Agar regex topa olmasa, standart usul bilan tozalab ko'ramiz
    clean_text = raw_response.strip()
    clean_text = re.sub(r'^```(?:json)?\s*', '', clean_text, flags=re.IGNORECASE)
    clean_text = re.sub(r'\s*```$', '', clean_text, flags=re.IGNORECASE)
    
    start_idx = clean_text.find('[')
    end_idx = clean_text.rfind(']')
    if start_idx != -1 and end_idx != -1:
        return json.loads(clean_text[start_idx:end_idx + 1])

    return json.loads(clean_text)

def _fetch_image_sync(keyword: str):
    """Unsplash'dan rasm yuklab olish"""
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        url = f"https://source.unsplash.com/800x600/?{encoded_keyword}"
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

        # 1-Slayd: Titul
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

        # Slayd sarlavhasi
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"{i+1}-Slayd")
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Asosiy matn
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
        elif isinstance(points, str):
            p = tf_content.paragraphs[0]
            p.text = f"• {points}"
            p.font.size = Pt(22)
            p.font.color.rgb = TEXT_COLOR

        # Rasm
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
