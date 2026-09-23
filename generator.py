import json
import os
import io
import re
import asyncio
import urllib.request
import urllib.parse
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

def _get_active_gemini_models():
    """Google'dan hozirgi faol modellar ro'yxatini olish"""
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={GEMINI_API_KEY}"
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            valid_models = []
            for m in data.get("models", []):
                name = m.get("name", "").replace("models/", "")
                methods = m.get("supportedGenerationMethods", [])
                if "generateContent" in methods and "flash" in name:
                    valid_models.append(name)
            if valid_models:
                return valid_models
    except Exception as e:
        print(f"[GENERATOR LOG] Modellarni olishda xato: {e}")
    
    return ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-2.5-flash"]

def _get_gemini_response_sync(prompt: str) -> str:
    if not GEMINI_API_KEY:
        raise Exception("GEMINI_API_KEY topilmadi!")

    models = _get_active_gemini_models()
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json"}
    }

    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GEMINI_API_KEY}"
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=15)
            if response.status_code == 200:
                data = response.json()
                text = data['candidates'][0]['content']['parts'][0]['text']
                return text
        except Exception as err:
            print(f"[GENERATOR LOG] Gemini {model} ulanish xatosi: {err}")
            continue

    raise Exception("Barcha Gemini modellarida xatolik yuz berdi.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    # SIZ aytgan g'oya: Gemini'dan rasm uchun o'ta batafsil va mos Promt so'raymiz
    prompt = f"""
    Mavzu: {topic}
    {f'Ssenariy: {user_script}' if user_script else ''}

    Vazifa: EXACTLY 6 ta slayd yaratib ber.
    
    JUDA MUHIM SHART:
    Har bir slayd uchun "image_prompt" maydonida shu slayd mazmunini AKSI ETTIRUVCHI, AI rasm generatori (Midjourney/Flux) tushunadigan, INGLIZ TILIDA O'TA BATAFSIL 1 ta tasviriy ko'rsatma (promt) yozib ber.

    Misollar:
    - Slayd suvni tejash haqida bo'lsa -> "A realistic photo of a clean water tap with fresh water drop, saving water concept, professional lighting, 4k"
    - Slayd chiqindini saralash haqida bo'lsa -> "A realistic photo of colorful waste sorting recycle bins with plastic and paper icons, clean environment, 4k"
    - Slayd yashil hududlar haqida bo'lsa -> "A realistic photo of people planting green tree saplings in a sunny city park, 4k"

    Javobingiz faqat va faqat quyidagi JSON formatida bo'lsin:
    {{
      "slides": [
        {{
          "title": "Slayd sarlavhasi",
          "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
          "image_prompt": "A realistic detailed prompt in english for AI image generator"
        }}
      ]
    }}
    """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = raw_response.strip()
        clean_text = re.sub(r"^```(?:json)?\s*", "", clean_text, flags=re.IGNORECASE)
        clean_text = re.sub(r"\s*```$", "", clean_text, flags=re.IGNORECASE).strip()

        data = json.loads(clean_text)
        if isinstance(data, dict) and "slides" in data:
            return data["slides"]
        elif isinstance(data, list):
            return data
    except Exception as e:
        print(f"[PARSER ERROR] JSON xatosi: {e}")

    return [
        {
            "title": topic,
            "content": ["Kirish va umumiy tushunchalar", "Asosiy yo'nalishlar", "Xulosalar"],
            "image_prompt": f"A professional high quality presentation banner about {topic}"
        }
    ]

def _fetch_image_sync(image_prompt: str, slide_index: int):
    """Gemini yaratgan batafsil promt bo'yicha sun'iy intellekt orqali mos rasm chizdirish"""
    if not image_prompt:
        image_prompt = "professional business presentation background"

    # URL uchun xavfsiz formatga keltirish
    encoded_prompt = urllib.parse.quote(image_prompt)
    
    # Pollinations AI (Flux modelidan foydalanib, o'ta aniq va sifatli rasm chizib beradi)
    url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=600&height=450&nologo=true&seed={slide_index + 42}"

    headers = {'User-Agent': 'Mozilla/5.0'}

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=12) as resp:
            if resp.status == 200:
                image_bytes = resp.read()
                # Rasm hajmi va sifati nazorati (Telegram limitidan oshmasligi uchun)
                if 2000 < len(image_bytes) < 4000000:
                    return io.BytesIO(image_bytes)
    except Exception as e:
        print(f"[IMAGE GENERATION ERROR] Rasm chizishda xatolik: {e}")

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

        # Muqova
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

        # Sarlavha
        title_text = slide_info.get("title", f"{i+1}-Slayd")
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Gemini bergan batafsil Promt orqali AI rasm yaratadi
        img_prompt = slide_info.get("image_prompt", "")
        img_stream = _fetch_image_sync(img_prompt, i)

        content_width = Inches(6.8) if img_stream else Inches(11.7)

        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), content_width, Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        if isinstance(points, list):
            for idx, point in enumerate(points):
                p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
                p.text = f"• {point}"
                p.font.size = Pt(20)
                p.font.color.rgb = TEXT_COLOR
                p.space_after = Pt(14)

        if img_stream:
            try:
                slide.shapes.add_picture(
                    img_stream, 
                    left=Inches(8.0), 
                    top=Inches(1.8), 
                    width=Inches(4.5)
                )
            except Exception as img_err:
                print(f"[IMAGE ERROR] Slaydga rasm qo'shishda xatolik: {img_err}")

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
