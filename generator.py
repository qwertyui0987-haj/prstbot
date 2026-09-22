import json
import os
import io
import asyncio
import urllib.request
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

def _get_gemini_response_sync(prompt: str) -> str:
    """Gemini API bilan muloqot qilish"""
    # So'rov uchun barqaror model
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text

async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}

        Vazifa: Berilgan ssenariy bo'yicha 6 ta slayd tayyorla.
        Matnlar to'liq va mazmunli o'zbek tilida bo'lsin.
        Javobni FAQAT quyidagi JSON formatida ber:
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["To'liq ma'lumotli punkt 1", "To'liq ma'lumotli punkt 2", "To'liq ma'lumotli punkt 3"]
          }}
        ]
        """
    else:
        prompt = f"""
        Mavzu: {topic}

        Vazifa: Ushbu mavzu bo'yicha 6 ta batafsil va mazmunli slayd yarat. 
        Har bir slayd uchun chuqur mazmunli 3-4 ta punkt yoz.
        Barcha matnlar o'zbek tilida bo'lsin.
        Javobni FAQAT toza JSON formatida ber (hech qanday qo'shimcha matnlarsiz):
        [
          {{
            "title": "Slayd sarlavhasi",
            "content": ["Atrof-muhitni muhofaza qilishning ahamiyati...", "Insoniyat faoliyatining tabiatga ta'siri...", "Tadbirlar va ko'kalamzorlashtirish..."]
          }}
        ]
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = raw_response.strip()
        
        # Markdown belgilarini tozalash
        if "```json" in clean_text:
            clean_text = clean_text.split("```json")[1].split("```")[0].strip()
        elif "```" in clean_text:
            clean_text = clean_text.split("```")[1].split("```")[0].strip()
            
        slides_data = json.loads(clean_text)
        return slides_data
    except Exception as e:
        print(f"!!! GEMINI API YOKI JSON XATOSI: {e} !!!")
        # Xatolik yuz bersa ham tushunarliroq va uzunroq ma'lumotlar
        return [
            {
                "title": f"{topic} haqida umumiy tushuncha",
                "content": [
                    "Ekologiya va atrof-muhit xavfsizligi bugungi kunda eng dolzarb masalalardan biridir.",
                    "Tabiiy resurslardan oqilona foydalanish kelajak avlod uchun juda muhim hisoblanadi.",
                    "Sanoat va maishiy chiqindilarni qayta ishlash darajasini oshirish lozim."
                ]
            },
            {
                "title": "Asosiy Muammolar va Ularning Oqibatlari",
                "content": [
                    "Havo hamda suv havzalarining sanoat chiqindilari bilan ifloslanishi.",
                    "Iqlim o'zgarishi va global issiqlik muammosining ortib borishi.",
                    "O'simlik va hayvonot dunyosidagi bioxilma-xillikning kamayishi."
                ]
            },
            {
                "title": "Amaliy Yechimlar va Xulosa",
                "content": [
                    "Yashil energiyaga (quyosh, shamol) o'tishni jadallashtirish.",
                    "Aholi o'rtasida ekologik madaniyatni oshirish hamda ko'kalamzorlashtirish.",
                    "Qayta tiklanuvchi manbalardan va ekologik toza texnologiyalardan foydalanish."
                ]
            }
        ]

def _fetch_image_sync():
    """Slaydlar uchun sifatli va barqaror rasm yuklash"""
    try:
        url = "https://picsum.photos/800/600"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as resp:
            return io.BytesIO(resp.read())
    except Exception as e:
        print(f"Rasm yuklashda xato: {e}")
        return None

def _build_pptx_sync(slides_data: list, output_filename: str) -> str:
    """PPTX faylini yaratish va shakllantirish"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    PRIMARY_COLOR = RGBColor(15, 32, 67)
    ACCENT_COLOR = RGBColor(0, 122, 255)
    TEXT_COLOR = RGBColor(50, 50, 50)

    for i, slide_info in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # 1-Slayd: Muqova
        if i == 0:
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.0))
            tf = title_box.text_frame
            tf.word_wrap = True
            
            p = tf.paragraphs[0]
            p.text = slide_info.get("title", "Prezentatsiya")
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            p.alignment = PP_ALIGN.CENTER
            
            p2 = tf.add_paragraph()
            p2.text = "\nTayyorladi: Sun'iy Intellekt (Gemini AI)"
            p2.font.size = Pt(20)
            p2.font.color.rgb = ACCENT_COLOR
            p2.alignment = PP_ALIGN.CENTER
            continue

        # Boshqa slaydlar
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = slide_info.get("title", f"Slayd {i+1}")
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        # Matn (Chap tomonda)
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        points = slide_info.get("content", [])
        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(12)

        # Rasm (O'ng tomonda)
        img_stream = _fetch_image_sync()
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(8.0), Inches(1.8), width=Inches(4.5))
            except Exception as e:
                print(f"Rasm joylashda xatolik: {e}")

    prs.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)
