import json
import os
import io
import re
import time
import asyncio
import urllib.request
import urllib.parse
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

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

    base_url = "https://generativelanguage.googleapis.com/v1beta/models"

    for model in models:
        endpoint = f"{base_url}/{model}:generateContent?key={GEMINI_API_KEY}"
        try:
            response = requests.post(endpoint, json=payload, headers=headers, timeout=25)
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
    """Gemini'dan matn va SAHIFAGA MOS ANIQ 1-3 TA INGLIZCHA REAL OBYEKT SO'RAYMIZ"""
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Ushbu ssenariy bo'yicha EXACTLY 6 ta slayd yaratib ber.
        
        JUDA MUHIM: Har bir slayd uchun shu slayd mazmunini aks ettiruvchi REAL va ANIQ inglizcha obyekt yoki sahna nomini yoz (image_keyword).
        Mavhum so'zlar yozma (masalan: "future", "strategy", "analysis", "examples" TAQIQLANADI).
        Faqat real ob'ekt yoki ko'rish mumkin bo'lgan sahna bo'lsin!
        Misollar:
        - Suv tejamkorligi -> "water tap drop"
        - Quyosh energiyasi -> "solar panel field"
        - Sun'iy intellekt -> "modern server room"
        - Moliya/Biznes -> "office team meeting"
        - Dehqonchilik -> "green wheat field"
        
        Javobingiz faqat va faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "water tap drop"
            }}
          ]
        }}
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha EXACTLY 6 ta slaydli prezentatsiya yarat.
        
        JUDA MUHIM: Har bir slayd uchun shu slayd mazmunini aks ettiruvchi REAL va ANIQ inglizcha obyekt yoki sahna nomini yoz (image_keyword).
        Mavhum so'zlar yozma (masalan: "future", "strategy", "analysis", "examples" TAQIQLANADI).
        Faqat real ob'ekt yoki ko'rish mumkin bo'lgan sahna bo'lsin!
        Misollar:
        - Suv tejamkorligi -> "water tap drop"
        - Quyosh energiyasi -> "solar panel field"
        - Sun'iy intellekt -> "modern server room"
        - Moliya/Biznes -> "office team meeting"

        Javobingiz faqat va faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "office team meeting"
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
        clean_text = re.sub(r"^\x60{3}(?:json)?\s*", "", clean_text, flags=re.IGNORECASE)
        clean_text = re.sub(r"\s*\x60{3}$", "", clean_text, flags=re.IGNORECASE).strip()

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
            "content": [f"{topic} - Kirish va umumiy tushunchalar", "Asosiy ustuvor yo'nalishlar", "Amaliy natijalar va xulosalar"],
            "image_keyword": "office team meeting"
        }
    ]

def _fetch_from_wikimedia(keyword: str):
    """Wikimedia Commons orqali mavzuga o'ta mos va keraksiz reklamalarsiz fotolarni yuklash"""
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        url = (
            f"https://commons.wikimedia.org/w/api.php?"
            f"action=query&generator=search&gsrsearch={encoded_keyword}&gsrlimit=8"
            f"&gsrnamespace=6&prop=imageinfo&iiprop=url|mime&format=json"
        )
        headers = {'User-Agent': 'TelegramSlideBot/1.0 (contact@telegram.org)'}
        req = urllib.request.Request(url, headers=headers)
        
        exclude_keywords = [
            'poster', 'flyer', 'advertisement', 'logo', 'banner', 'signboard', 
            'infographic', 'card', 'diagram', 'map', 'flag', 'icon', 'symbol', 'vector'
        ]

        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            pages = data.get('query', {}).get('pages', {})
            for page_id, page in pages.items():
                title_lower = page.get('title', '').lower()
                
                if any(bad_word in title_lower for bad_word in exclude_keywords):
                    continue

                imageinfo = page.get('imageinfo', [{}])[0]
                mime = imageinfo.get('mime', '')
                img_url = imageinfo.get('url', '')

                if mime in ['image/jpeg', 'image/png'] and img_url:
                    img_req = urllib.request.Request(img_url, headers=headers)
                    with urllib.request.urlopen(img_req, timeout=10) as img_resp:
                        content = img_resp.read()
                        if 5000 < len(content) < 4000000:
                            return io.BytesIO(content)
    except Exception as e:
        print(f"[IMAGE LOG] Wikimedia xatosi ({keyword}): {e}")
    return None

def _fetch_from_pollinations_flux(keyword: str, slide_index: int):
    """Pollinations AI orqali har qanday mavzuga 100% mos va aniq fotolarni chizdirish"""
    try:
        if not keyword:
            keyword = "modern office meeting"
            
        detailed_prompt = (
            f"A high quality realistic professional photo of {keyword}, "
            f"clean background, studio lighting, detailed visual, nologo, no text"
        )
        encoded_prompt = urllib.parse.quote(detailed_prompt)
        
        url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=800&height=600&nologo=true&seed={slide_index + 100}"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=12) as resp:
            if resp.status == 200:
                content = resp.read()
                if 3000 < len(content) < 4000000:
                    return io.BytesIO(content)
    except Exception as e:
        print(f"[IMAGE LOG] Pollinations Flux xatosi ({keyword}): {e}")
    return None

def _fetch_image_sync(keyword: str, slide_index: int):
    """Slayd kalit so'ziga mos fotolarni qidirish va yuklash"""
    if not keyword:
        keyword = "office team meeting"

    clean_keyword = re.sub(r'[^a-zA-Z0-9\s]', '', keyword).strip()
    if not clean_keyword:
        clean_keyword = "office team meeting"

    img_stream = _fetch_from_wikimedia(clean_keyword)
    if img_stream:
        return img_stream

    time.sleep(1)

    img_stream = _fetch_from_pollinations_flux(clean_keyword, slide_index)
    if img_stream:
        return img_stream

    return None

def _build_pptx_sync(slides_data: list, output_filename: str) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Zamonaviy ranglar palitrasi
    PRIMARY_COLOR = RGBColor(16, 37, 66)      # To'q ko'k / Navy Blue
    ACCENT_COLOR = RGBColor(0, 150, 214)     # Och moviy / Cyan Blue
    TEXT_COLOR = RGBColor(45, 55, 72)        # Zamonaviy to'q kulrang
    MUTED_LINE_COLOR = RGBColor(226, 232, 240) # Mayin kulrang chiziq

    for i, slide_info in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        title_text = slide_info.get("title", f"{i+1}-Slayd")

        # Content matnlarini normalize qilish
        raw_content = slide_info.get("content") or slide_info.get("points") or []
        if isinstance(raw_content, str):
            points = [p.strip() for p in raw_content.split("\n") if p.strip()]
        elif isinstance(raw_content, list):
            points = [str(p).strip() for p in raw_content if str(p).strip()]
        else:
            points = []

        if not points:
            points = [f"{title_text} bo'yicha asosiy tushunchalar", "Tahlil va amaliy ko'rsatkichlar"]

        # -------------------------------------------------------------
        # 1-Slayd: Muqova (Title + Subtitle + Dizayn Elementlari)
        # -------------------------------------------------------------
        if i == 0:
            # Yuqori brending hoshiyasi (Header Bar)
            top_bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.25)
            )
            top_bar.fill.solid()
            top_bar.fill.fore_color.rgb = PRIMARY_COLOR
            top_bar.line.fill.background()

            # Pastki brending hoshiyasi (Footer Bar)
            bottom_bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.3), Inches(13.333), Inches(0.2)
            )
            bottom_bar.fill.solid()
            bottom_bar.fill.fore_color.rgb = ACCENT_COLOR
            bottom_bar.line.fill.background()

            # Sarlavha matn qutisi
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(1.8))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(42)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            p.alignment = PP_ALIGN.CENTER

            # Sarlavha ostidagi zamonaviy ajratuvchi chiziq (Accent line)
            divider = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.666), Inches(3.8), Inches(2.0), Inches(0.06)
            )
            divider.fill.solid()
            divider.fill.fore_color.rgb = ACCENT_COLOR
            divider.line.fill.background()

            # Muqova osti matni (Subtitle)
            sub_box = slide.shapes.add_textbox(Inches(1.5), Inches(4.3), Inches(10.333), Inches(2.5))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            for idx, pt in enumerate(points):
                p_sub = tf_sub.add_paragraph() if idx > 0 else tf_sub.paragraphs[0]
                p_sub.text = pt
                p_sub.font.size = Pt(20)
                p_sub.font.color.rgb = TEXT_COLOR
                p_sub.alignment = PP_ALIGN.CENTER
                p_sub.space_after = Pt(8)
            continue

        # -------------------------------------------------------------
        # 2-6 Slaydlar: Ichki Slaydlar Dizayni
        # -------------------------------------------------------------
        # 1. Yuqori zamonaviy hoshiya (Header Bar)
        top_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = PRIMARY_COLOR
        top_bar.line.fill.background()

        # 2. Sarlavha
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.9))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(30)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # 3. Sarlavha ostidagi vizual chiziq (Header line)
        header_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.7), Inches(0.03)
        )
        header_line.fill.solid()
        header_line.fill.fore_color.rgb = MUTED_LINE_COLOR
        header_line.line.fill.background()

        # Slayd mazmuniga mos rasm olish
        keyword = slide_info.get("image_keyword", "")
        img_stream = _fetch_image_sync(keyword, i)

        if img_stream:
            content_width = Inches(6.8)
        else:
            content_width = Inches(11.7)

        # 4. Matn qutisi
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), content_width, Inches(4.8))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(19)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(14)

        # 5. Rasmni joylashtirish
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

        # 6. Pastki burchak dizayn urg'usi (Footer Accent Strip)
        footer_stripe = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.1), Inches(1.5), Inches(0.05)
        )
        footer_stripe.fill.solid()
        footer_stripe.fill.fore_color.rgb = ACCENT_COLOR
        footer_stripe.line.fill.background()

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
