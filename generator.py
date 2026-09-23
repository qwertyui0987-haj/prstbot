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
    """Google'dan hozirgi faol modellar ro'yxatini dinamik ravishda olish"""
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
    """Gemini API orqali so'rov yuborish"""
    if not GEMINI_API_KEY:
        raise Exception("GEMINI_API_KEY topilmadi! Railway Variables bo'limiga GEMINI_API_KEY ni qo'shing.")

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
            else:
                print(f"[GENERATOR LOG] Gemini {model} status: {response.status_code}, msg: {response.text}")
        except Exception as err:
            print(f"[GENERATOR LOG] Gemini {model} ulanish xatosi: {err}")
            continue

    raise Exception("Barcha Gemini modellarida xatolik yuz berdi.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Ushbu ssenariy bo'yicha EXACTLY 6 ta slayd yaratib ber.
        MUHIM: Har bir slaydning MAZMUNIGA va YAZILGAN FIKRLARIGA to'g'ri keladigan, ingliz tilidagi O'TA ANIQ 1-2 ta visual kalit so'z yoz (image_keyword).
        Misol uchun: agar slayd suvni tejash va kranni yopish haqida bo'lsa -> "water faucet" yoki "water drop".
        
        Javobingiz faqat va faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "water faucet"
            }}
          ]
        }}
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha EXACTLY 6 ta slaydli prezentatsiya yarat.
        MUHIM: Har bir slaydning MAZMUNIGA va YAZILGAN FIKRLARIGA to'g'ri keladigan, ingliz tilidagi O'TA ANIQ 1-2 ta visual kalit so'z yoz (image_keyword).
        Misol uchun: agar slayd quyosh energiyasi haqida bo'lsa -> "solar panel", agar kompyuter haqida bo'lsa -> "laptop keyboard".

        Javobingiz faqat va faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "solar panel"
            }}
          ]
        }}
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
    except Exception as e:
        print(f"[GENERATOR ERROR] API xatosi: {e}")
        raw_response = ""

    try:
        clean_text = raw_response.strip()
        clean_text = re.sub(r"^```(?:json)?\s*", "", clean_text, flags=re.IGNORECASE)
        clean_text = re.sub(r"\s*```$", "", clean_text, flags=re.IGNORECASE).strip()

        data = json.loads(clean_text)
        if isinstance(data, dict) and "slides" in data and isinstance(data["slides"], list):
            return data["slides"]
        elif isinstance(data, list) and len(data) > 0:
            return data
    except Exception as e:
        print(f"[PARSER ERROR] JSON o'qishda xatolik: {e}")

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

def _fetch_image_sync(keyword: str, slide_title: str):
    """Mavzuga va slayd mazmuniga aniq mos rasm yuklab olish"""
    if not keyword:
        keyword = slide_title if slide_title else "presentation"

    # Kalit so'zni tozalaymiz va probellarni URL formatiga o'tkazamiz
    clean_keyword = re.sub(r'[^a-zA-Z0-9\s]', '', keyword).strip()
    encoded_keyword = urllib.parse.quote(clean_keyword)
    
    # Aniq tasvir yaratuvchi / qidiruvchi manbalar (Realistic photo yo'nalishida)
    urls = [
        f"https://image.pollinations.ai/prompt/realistic%20photo%20of%20{encoded_keyword}%20clean%20professional%20style?width=800&height=600&nologo=true",
        f"https://loremflickr.com/800/600/{encoded_keyword}",
        f"https://picsum.photos/seed/{encoded_keyword}/800/600"
    ]

    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

    for url in urls:
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=8) as resp:
                if resp.status == 200:
                    image_bytes = resp.read()
                    if len(image_bytes) > 3000: # To'liq rasm ekanligiga ishonch hosil qilish
                        return io.BytesIO(image_bytes)
        except Exception as e:
            print(f"[IMAGE LOG] URL bo'yicha rasm yuklashda xato ({url}): {e}")
            continue

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

        # 1-Slayd: Muqova
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

        # Slayd mazmuniga mos rasm olish
        keyword = slide_info.get("image_keyword", "")
        img_stream = _fetch_image_sync(keyword, title_text)

        if img_stream:
            content_width = Inches(6.8)
        else:
            content_width = Inches(11.7)

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
        else:
            p = tf_content.paragraphs[0]
            p.text = f"• {points}"
            p.font.size = Pt(20)
            p.font.color.rgb = TEXT_COLOR

        # Rasmni joylashtirish
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
