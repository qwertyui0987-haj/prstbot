import json
import os
import io
import re
import asyncio
import urllib.request
import urllib.parse
from groq import Groq
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

GROQ_API_KEY = os.getenv("GROQ_API_KEY")

def _get_groq_response_sync(prompt: str) -> str:
    """Groq API orqali Llama-3 modeliga so'rov yuborish"""
    if not GROQ_API_KEY:
        raise Exception("GROQ_API_KEY topilmadi! Railway Variables bo'limiga GROQ_API_KEY ni qo'shing.")

    client = Groq(api_key=GROQ_API_KEY)

    response = client.chat.completions.create(
        messages=[
            {
                "role": "system",
                "content": "Siz faqat va faqat standart JSON formatida javob beradigan yordamchisiz. Boshqa hech qanday kirish yoki chiqish matnlari yozmang."
            },
            {
                "role": "user",
                "content": prompt,
            }
        ],
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        response_format={"type": "json_object"}
    )

    if response.choices and len(response.choices) > 0:
        return response.choices[0].message.content

    raise Exception("Groq API dan bo'sh javob qaytdi.")

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Ushbu ssenariy bo'yicha EXACTLY 6 ta slayd yaratib ber.
        Javobingiz faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "business"
            }}
          ]
        }}
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha EXACTLY 6 ta slaydli prezentatsiya yarat.
        Javobingiz faqat quyidagi JSON tuzilmasida bo'lishi shart:
        {{
          "slides": [
            {{
              "title": "Slayd sarlavhasi",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "technology"
            }}
          ]
        }}
        """

    try:
        raw_response = await asyncio.to_thread(_get_groq_response_sync, prompt)
    except Exception as e:
        print(f"[GENERATOR ERROR] Groq API Xatosi: {e}")
        raw_response = ""

    # JSON obyektini parse qilish
    try:
        data = json.loads(raw_response)
        if isinstance(data, dict) and "slides" in data and isinstance(data["slides"], list):
            return data["slides"]
        elif isinstance(data, list) and len(data) > 0:
            return data
    except Exception as e:
        print(f"[PARSER ERROR] JSON ni o'qishda xatolik: {e}")

    # Agar API da kutilmagan xatolik bo'lsa, zaxira slaydlar
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
    return await asyncio-to_thread(_build_pptx_sync, slides_data, output_filename)
