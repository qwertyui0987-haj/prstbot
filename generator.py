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

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

def _get_gemini_response_sync(prompt: str) -> str:
    """Gemini API orqali so'rov yuborish"""
    if not GEMINI_API_KEY:
        raise Exception("GEMINI_API_KEY topilmadi! Railway Variables bo'limiga GEMINI_API_KEY ni qo'shing.")

    models = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json"}
    }

    base_url = "https://generativelanguage.googleapis.com/v1beta/models"

    for model in models:
        endpoint = f"{base_url}/{model}:generateContent?key={GEMINI_API_KEY}"
        try:
            response = requests.post(endpoint, json=payload, headers=headers, timeout=15)
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

def _get_default_slides(topic: str) -> list:
    """Xatolik yuz berganda to'liq 6 ta slayd qaytaruvchi zaxira funksiyasi"""
    return [
        {
            "title": topic,
            "content": ["Taqdimot va asosiy tushunchalar", "Mavzuning umumiy sharhi va tahlili"],
            "image_keyword": "office team meeting"
        },
        {
            "title": f"1. {topic} - Asosiy maqsad va mohiyat",
            "content": ["Mavzuning ustuvor yo'nalishlari", "Asosiy vazifalar va maqsadlar", "Kutilayotgan amaliy natijalar"],
            "image_keyword": "business strategy chart"
        },
        {
            "title": "2. Jarayon va amaliy tahlil",
            "content": ["Tizimli yondashuv va tahlil", "Amaliyotdagi holat va ko'rsatkichlar", "Mavjud imkoniyatlar sharhi"],
            "image_keyword": "financial chart on laptop"
        },
        {
            "title": "3. Texnologiya va innovatsiyalar",
            "content": ["Zamonaviy yechimlar va vositalar", "Samadorlikni oshirish usullari", "Avtomatlashtirish bosqichlari"],
            "image_keyword": "modern server room"
        },
        {
            "title": "4. Amaliy natijalar va misollar",
            "content": ["Muvaffaqiyatli tajribalar", "Resurslardan unumli foydalanish", "Olingan xulosalar va tajriba"],
            "image_keyword": "drip irrigation system"
        },
        {
            "title": "5. Xulosa va kelajak rejalari",
            "content": ["Istiqboldagi ustuvor vazifalar", "Rivojlanish bosqichlari", "Yakuniy xulosalar"],
            "image_keyword": "robotics lab engineer"
        }
    ]

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    instructions = """
    JUDA MUHIM - `image_keyword` QOIDASI:
    Har bir slayd uchun 100% aniq va sifatli foto beradigan real obyekt, buyum yoki sahnani ingliz tilida (1-3 so'z) yoz.
    HECH QACHON mavhum yoki umumiy so'z ishlatma (masalan: "examples", "future", "strategy", "success", "analysis", "concept", "business", "presentation").
    
    To'g'ri misollar:
    - Suv va kran -> "water tap"
    - Boshqaruv / Majlis -> "office team meeting"
    - Moliya / Tahlil -> "financial chart on laptop"
    - Texnologiya / AI -> "modern server room"
    - Qishloq xo'jaligi -> "drip irrigation system"
    - Kelajak / Innovatsiya -> "robotics lab engineer"
    """

    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Ushbu ssenariy bo'yicha EXACTLY 6 ta slayd yaratib ber.
        
        {instructions}

        Javobingiz faqat va faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "water tap"
            }}
          ]
        }}
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha EXACTLY 6 ta slaydli prezentatsiya yarat.
        
        {instructions}

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
        clean_text = raw_response.strip()
        clean_text = re.sub(r"^\x60{3}(?:json)?\s*", "", clean_text, flags=re.IGNORECASE)
        clean_text = re.sub(r"\s*\x60{3}$", "", clean_text, flags=re.IGNORECASE).strip()

        data = json.loads(clean_text)
        slides = []
        if isinstance(data, dict) and "slides" in data and isinstance(data["slides"], list):
            slides = data["slides"]
        elif isinstance(data, list) and len(data) > 0:
            slides = data

        if slides and len(slides) >= 3:
            return slides
    except Exception as e:
        print(f"[PARSER ERROR] JSON o'qishda xatolik: {e}")

    return _get_default_slides(topic)

def _fetch_from_wikimedia(keyword: str):
    """Wikimedia Commons orqali mavzuga o'ta mos fotolarni yuklash"""
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        url = (
            f"https://commons.wikimedia.org/w/api.php?"
            f"action=query&generator=search&gsrsearch={encoded_keyword}&gsrlimit=8"
            f"&gsrnamespace=6&prop=imageinfo&iiprop=url|mime&format=json"
        )
        headers = {'User-Agent': 'TelegramSlideBot/1.0 (contact@telegram.org)'}
        req = urllib.request.Request(url, headers=headers)
        
        exclude_keywords = ['poster', 'flyer', 'advertisement', 'logo', 'banner', 'signboard', 'infographic', 'card']

        with urllib.request.urlopen(req, timeout=6) as resp:
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
                    with urllib.request.urlopen(img_req, timeout=8) as img_resp:
                        content = img_resp.read()
                        if len(content) > 5000:
                            return io.BytesIO(content)
    except Exception as e:
        print(f"[IMAGE LOG] Wikimedia xatosi ({keyword}): {e}")
    return None

def _fetch_from_pollinations(keyword: str, slide_index: int):
    """Pollinations AI orqali rasm olish"""
    try:
        encoded_keyword = urllib.parse.quote(f"high quality photo of {keyword}")
        url = f"https://image.pollinations.ai/prompt/{encoded_keyword}?width=800&height=600&nologo=true&seed={slide_index + 10}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            if resp.status == 200:
                content = resp.read()
                if len(content) > 3000:
                    return io.BytesIO(content)
    except Exception as e:
        print(f"[IMAGE LOG] Pollinations xatosi ({keyword}): {e}")
    return None

def _fetch_image_sync(keyword: str, slide_index: int):
    """Slayd kalit so'ziga mos rasmlarni yuklash"""
    if not keyword:
        keyword = "office team meeting"

    clean_keyword = re.sub(r'[^a-zA-Z0-9\s]', '', keyword).strip()
    if not clean_keyword:
        clean_keyword = "office team meeting"

    img_stream = _fetch_from_wikimedia(clean_keyword)
    if img_stream:
        return img_stream

    time.sleep(1)

    img_stream = _fetch_from_pollinations(clean_keyword, slide_index)
    if img_stream:
        return img_stream

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

        # 1-Slayd: Muqova (Title + Subtitle)
        if i == 0:
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(1.8))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(44)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            p.alignment = PP_ALIGN.CENTER

            # Muqova osti matni (Subtitle)
            sub_box = slide.shapes.add_textbox(Inches(1.5), Inches(4.2), Inches(10.333), Inches(2.5))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            for idx, pt in enumerate(points):
                p_sub = tf_sub.add_paragraph() if idx > 0 else tf_sub.paragraphs[0]
                p_sub.text = pt
                p_sub.font.size = Pt(22)
                p_sub.font.color.rgb = TEXT_COLOR
                p_sub.alignment = PP_ALIGN.CENTER
                p_sub.space_after = Pt(10)
            continue

        # Keyingi slaydlarning Sarlavhasi
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Rasm va Matn joylashuvi
        keyword = slide_info.get("image_keyword", "")
        img_stream = _fetch_image_sync(keyword, i)

        if img_stream:
            content_width = Inches(6.8)
        else:
            content_width = Inches(11.7)

        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), content_width, Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(20)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(14)

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
