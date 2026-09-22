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
    # Model va maxsus JSON rejimini yoqish
    model = genai.GenerativeModel(
        "gemini-1.5-flash",
        generation_config={"response_mime_type": "application/json"}
    )
    response = model.generate_content(prompt)
    return response.text

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Berilgan ssenariydan foydalanib 5-6 ta slayd tayyorla.
        Javob FAQAT quyidagi JSON ro'yxati (array) bo'lishi shart:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["To'liq ma'lumot 1", "To'liq ma'lumot 2", "To'liq ma'lumot 3"],
            "image_keyword": "nature"
          }}
        ]
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Ushbu mavzu bo'yicha 5-6 ta batafsil slayd tayyorla.
        Javob FAQAT quyidagi JSON ro'yxati (array) bo'lishi shart:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Atrof-muhitni muhofaza qilish...", "Insoniyat faoliyatining ta'siri...", "Chiqindilarni qayta ishlash..."],
            "image_keyword": "environment"
          }}
        ]
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        return json.loads(raw_response)
    except Exception as e:
        print(f"!!! GEMINI API XATOSI: {e} !!!")
        # Zaxira uchun kamida 5 ta to'liq slayd!
        return [
            {
                "title": topic,
                "content": [
                    "Ushbu mavzu bugungi kunda jamiyatimizda muhim o'rin tutadi.",
                    "Asosiy tushunchalar va muammoning mohiyati ko'rib chiqiladi."
                ],
                "image_keyword": "nature"
            },
            {
                "title": "Mavzuning Dolzarbligi",
                "content": [
                    "Zamonaviy texnologiyalar va atrof-muhit o'rtasidagi muvozanat.",
                    "So'nggi yillarda kuzatilayotgan asosiy o'zgarishlar va ko'rsatkichlar.",
                    "Inson salomatligi va xavfsizligiga ta'siri."
                ],
                "image_keyword": "earth"
            },
            {
                "title": "Asosiy Muammolar va Tahlil",
                "content": [
                    "Resurslardan unumsiz foydalanish va chiqindilar muammosi.",
                    "Sanoat rivojlanishining salbiy oqibatlari.",
                    "Ekotizim barqarorligini saqlashdagi qiyinchiliklar."
                ],
                "image_keyword": "pollution"
            },
            {
                "title": "Tavsiya va Yechimlar",
                "content": [
                    "Muammoni hal etishga qaratilgan amaliy va samarali qadamlar.",
                    "Zamonaviy yashil texnologiyalarni tatbiq etish.",
                    "Aholi va yoshlar o'rtasida tushuntirish ishlarini olib borish."
                ],
                "image_keyword": "green energy"
            },
            {
                "title": "Xulosa",
                "content": [
                    "Barcha ko'rib chiqilgan masalalarning umumiy xulosasi.",
                    "Kelajakdagi maqsad va istiqbolli rejalarni belgilab olish."
                ],
                "image_keyword": "success"
            }
        ]

def _fetch_image_sync(keyword: str):
    """Mavzuga mos rasmlarni Unsplash xizmatidan olish"""
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

        # 1-Slayd: Titul (Hech qanday AI yozuvlarisiz)
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

        # Sarlavha (Katta va aniq)
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"{i+1}-Slayd")
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Asosiy matn (Tushunarli va 22pt o'lchamda)
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

        # Rasm (O'ng tomonda)
        keyword = slide_info.get("image_keyword", "nature")
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
