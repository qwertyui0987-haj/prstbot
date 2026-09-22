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
    """Modelleri sırayla deneyerek JSON formatında yanıt alma"""
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
            print(f"[GENERATOR] Model deneniyor: {model_name}")
            # JSON yanıt tipini zorunlu kılıyoruz
            model = genai.GenerativeModel(
                model_name,
                generation_config={"response_mime_type": "application/json"}
            )
            response = model.generate_content(prompt)
            if response and response.text:
                print(f"[GENERATOR] Başarıyla kullanılan model: {model_name}")
                return response.text
        except Exception as err:
            # Standart yapılandırma ile tekrar dene
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(prompt)
                if response and response.text:
                    return response.text
            except Exception:
                pass
            print(f"[GENERATOR] {model_name} hatası: {err}")
            last_exception = err

    if last_exception:
        raise last_exception
    raise Exception("Çalışan bir Gemini modeli bulunamadı.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    print(f"\n==========================================")
    print(f"[GENERATOR] Çağrı alındı! Konu: '{topic}'")
    print(f"==========================================\n")

    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Berilgan ssenariy bo'yicha 6 ta slayd tayyorla. 
        Javobingiz FAQAT va FAQAT toza JSON formatida bo'lsin:
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
        Javobingiz FAQAT va FAQAT toza JSON formatida bo'lsin:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Mavzuga oid batafsil fikr 1", "Mavzuga oid batafsil fikr 2", "Mavzuga oid batafsil fikr 3"],
            "image_keyword": "technology"
          }}
        ]
        """

    raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
    print(f"[GENERATOR] Yanıt alındı. Uzunluk: {len(raw_response)}")

    # Temizleme adımları
    clean_text = raw_response.strip()
    clean_text = re.sub(r'^```(?:json)?\s*', '', clean_text, flags=re.IGNORECASE)
    clean_text = re.sub(r'\s*```$', '', clean_text, flags=re.IGNORECASE)
    clean_text = clean_text.strip()

    # Yalnızca ilk '[' ile son ']' arasındaki geçerli JSON dizisini alıyoruz
    start_idx = clean_text.find('[')
    end_idx = clean_text.rfind(']')

    if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
        json_str = clean_text[start_idx:end_idx + 1]
        return json.loads(json_str)

    return json.loads(clean_text)

def _fetch_image_sync(keyword: str):
    """Görsel çekme işlevi"""
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

        # 1. Slayt: Başlık
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

        # Diğer slayt başlıkları
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"{i+1}-Slayd")
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # İçerik
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(7.0), Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(22)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(14)

        # Görsel
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
