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

from docx import Document
from docx.shared import Inches as DocxInches, Pt as DocxPt
from docx.enum.text import WD_ALIGN_PARAGRAPH

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

def _get_active_gemini_models():
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

    base_url = "https://generativelanguage.googleapis.com/v1beta/models"

    for model in models:
        endpoint = f"{base_url}/{model}:generateContent?key={GEMINI_API_KEY}"
        try:
            response = requests.post(endpoint, json=payload, headers=headers, timeout=25)
            if response.status_code == 200:
                data = response.json()
                return data['candidates'][0]['content']['parts'][0]['text']
        except Exception:
            continue

    raise Exception("Barcha Gemini modellarida xatolik yuz berdi.")

# ================= SLAYD GENERATSIYASI =================
async def generate_presentation_content(topic: str, user_script: str = None) -> list:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Ssenariy: {user_script}
        Vazifa: Ushbu ssenariy bo'yicha EXACTLY 6 ta slayd yaratib ber.
        Har bir slayd uchun real inglizcha image_keyword yoz.
        Javobingiz faqat JSON formatida bo'lsin:
        {{
          "slides": [
            {{
              "title": "Sarlavha",
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
        Real inglizcha image_keyword yoz.
        Javobingiz faqat JSON formatida bo'lsin:
        {{
          "slides": [
            {{
              "title": "Sarlavha",
              "content": ["Fikr 1", "Fikr 2", "Fikr 3"],
              "image_keyword": "office team meeting"
            }}
          ]
        }}
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = re.sub(r"^\x60{3}(?:json)?\s*", "", raw_response.strip(), flags=re.IGNORECASE)
        clean_text = re.sub(r"\s*\x60{3}$", "", clean_text, flags=re.IGNORECASE).strip()
        data = json.loads(clean_text)
        if isinstance(data, dict) and "slides" in data:
            return data["slides"]
    except Exception:
        pass

    return [{
        "title": topic,
        "content": [f"{topic} - Kirish", "Asosiy qism va tahlil", "Xulosalar"],
        "image_keyword": "office team meeting"
    }]

# ================= REFERAT GENERATSIYASI =================
async def generate_essay_content(topic: str, user_script: str = None) -> dict:
    if user_script:
        prompt = f"""
        Mavzu: {topic}
        Reja/Ssenariy: {user_script}
        Vazifa: Ushbu reja asosida EXACTLY 6 sahifalik akademik referat matnini yarat. 
        Har bir bob o'ta batafsil, ilmiy va kengaytirilgan 3-4 ta abzasdan iborat bo'lsin.
        Javob faqat JSON formatida bo'lsin:
        {{
          "title": "{topic}",
          "introduction": "Juda batafsil va kengaytirilgan kirish matni...",
          "body_sections": [
            {{"section_title": "1-bob nomi", "section_text": "O'ta batafsil va uzoq matn..."}},
            {{"section_title": "2-bob nomi", "section_text": "O'ta batafsil va uzoq matn..."}},
            {{"section_title": "3-bob nomi", "section_text": "O'ta batafsil va uzoq matn..."}}
          ],
          "conclusion": "Batafsil va to'liq xulosa matni...",
          "references": [
            "1. Muallif A. Ilmiy kitob yoki darslik nomi. - Toshkent, 2021.",
            "2. Muallif B. Maxsus tadqiqotlar va manbalar. - Toshkent, 2023.",
            "3. Rasmiy axborot va darslik manbalari."
          ]
        }}
        """
    else:
        prompt = f"""
        Mavzu: {topic}
        Vazifa: Ushbu mavzu bo'yicha EXACTLY 6 sahifalik akademik referat yarat.
        Har bir bob o'ta batafsil, ilmiy va kengaytirilgan 3-4 ta abzasdan iborat bo'lsin.
        Javob faqat JSON formatida bo'lsin:
        {{
          "title": "{topic}",
          "introduction": "Juda batafsil va kengaytirilgan kirish matni...",
          "body_sections": [
            {{"section_title": "1-bob nomi", "section_text": "O'ta batafsil va uzoq matn..."}},
            {{"section_title": "2-bob nomi", "section_text": "O'ta batafsil va uzoq matn..."}},
            {{"section_title": "3-bob nomi", "section_text": "O'ta batafsil va uzoq matn..."}}
          ],
          "conclusion": "Batafsil va to'liq xulosa matni...",
          "references": [
            "1. Muallif A. Ilmiy kitob yoki darslik nomi. - Toshkent, 2021.",
            "2. Muallif B. Maxsus tadqiqotlar va manbalar. - Toshkent, 2023.",
            "3. Rasmiy axborot va darslik manbalari."
          ]
        }}
        """

    try:
        raw_response = await asyncio.to_thread(_get_gemini_response_sync, prompt)
        clean_text = re.sub(r"^\x60{3}(?:json)?\s*", "", raw_response.strip(), flags=re.IGNORECASE)
        clean_text = re.sub(r"\s*\x60{3}$", "", clean_text, flags=re.IGNORECASE).strip()
        return json.loads(clean_text)
    except Exception:
        return {
            "title": topic,
            "introduction": f"{topic} bo'yicha kirish qismi va umumiy tushunchalar.",
            "body_sections": [{"section_title": "1-bob. Asosiy tushunchalar", "section_text": f"{topic} ning batafsil tahlili va ko'rsatkichlari."}],
            "conclusion": "Xulosa va amaliy takliflar.",
            "references": ["1. Ilmiy adabiyotlar va darsliklar ro'yxati."]
        }

def _fetch_from_wikimedia(keyword: str):
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        url = f"https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrsearch={encoded_keyword}&gsrlimit=5&gsrnamespace=6&prop=imageinfo&iiprop=url|mime&format=json"
        headers = {'User-Agent': 'TelegramSlideBot/1.0'}
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            pages = data.get('query', {}).get('pages', {})
            for page_id, page in pages.items():
                imageinfo = page.get('imageinfo', [{}])[0]
                mime = imageinfo.get('mime', '')
                img_url = imageinfo.get('url', '')
                if mime in ['image/jpeg', 'image/png'] and img_url:
                    img_req = urllib.request.Request(img_url, headers=headers)
                    with urllib.request.urlopen(img_req, timeout=10) as img_resp:
                        content = img_resp.read()
                        if 5000 < len(content) < 4000000:
                            return io.BytesIO(content)
    except Exception:
        pass
    return None

def _fetch_from_pollinations_flux(keyword: str, slide_index: int):
    try:
        if not keyword:
            keyword = "modern office meeting"
        detailed_prompt = f"A high quality realistic professional photo of {keyword}, clean background, studio lighting"
        encoded_prompt = urllib.parse.quote(detailed_prompt)
        url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=800&height=600&nologo=true&seed={slide_index + 100}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=12) as resp:
            if resp.status == 200:
                content = resp.read()
                if 3000 < len(content) < 4000000:
                    return io.BytesIO(content)
    except Exception:
        pass
    return None

def _fetch_image_sync(keyword: str, slide_index: int):
    if not keyword:
        keyword = "office team meeting"
    clean_keyword = re.sub(r'[^a-zA-Z0-9\s]', '', keyword).strip() or "office team meeting"
    img_stream = _fetch_from_wikimedia(clean_keyword)
    if img_stream:
        return img_stream
    time.sleep(1)
    return _fetch_from_pollinations_flux(clean_keyword, slide_index)

def _build_pptx_sync(slides_data: list, output_filename: str) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    PRIMARY_COLOR = RGBColor(16, 37, 66)
    ACCENT_COLOR = RGBColor(0, 150, 214)
    TEXT_COLOR = RGBColor(45, 55, 72)
    MUTED_LINE_COLOR = RGBColor(226, 232, 240)

    for i, slide_info in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        title_text = slide_info.get("title", f"{i+1}-Slayd")

        raw_content = slide_info.get("content") or slide_info.get("points") or []
        if isinstance(raw_content, str):
            points = [p.strip() for p in raw_content.split("\n") if p.strip()]
        elif isinstance(raw_content, list):
            points = [str(p).strip() for p in raw_content if str(p).strip()]
        else:
            points = [f"{title_text} bo'yicha asosiy tushunchalar"]

        if i == 0:
            top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.25))
            top_bar.fill.solid()
            top_bar.fill.fore_color.rgb = PRIMARY_COLOR
            top_bar.line.fill.background()

            bottom_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.3), Inches(13.333), Inches(0.2))
            bottom_bar.fill.solid()
            bottom_bar.fill.fore_color.rgb = ACCENT_COLOR
            bottom_bar.line.fill.background()

            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(1.8))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(42)
            p.font.bold = True
            p.font.color.rgb = PRIMARY_COLOR
            p.alignment = PP_ALIGN.CENTER

            divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.666), Inches(3.8), Inches(2.0), Inches(0.06))
            divider.fill.solid()
            divider.fill.fore_color.rgb = ACCENT_COLOR
            divider.line.fill.background()

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

        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = PRIMARY_COLOR
        top_bar.line.fill.background()

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.9))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(30)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_COLOR

        header_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.7), Inches(0.03))
        header_line.fill.solid()
        header_line.fill.fore_color.rgb = MUTED_LINE_COLOR
        header_line.line.fill.background()

        keyword = slide_info.get("image_keyword", "")
        img_stream = _fetch_image_sync(keyword, i)
        content_width = Inches(6.8) if img_stream else Inches(11.7)

        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), content_width, Inches(4.8))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for idx, point in enumerate(points):
            p = tf_content.add_paragraph() if idx > 0 else tf_content.paragraphs[0]
            p.text = f"• {point}"
            p.font.size = Pt(19)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(14)

        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, left=Inches(8.0), top=Inches(1.8), width=Inches(4.5))
            except Exception:
                pass

        footer_stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.1), Inches(1.5), Inches(0.05))
        footer_stripe.fill.solid()
        footer_stripe.fill.fore_color.rgb = ACCENT_COLOR
        footer_stripe.line.fill.background()

    prs.save(output_filename)
    return output_filename

def _build_docx_sync(essay_data: dict, output_filename: str) -> str:
    doc = Document()
    
    # ---------------------------------------------------------
    # PAGE 1: TITUL VARAG'I (STANDART)
    # ---------------------------------------------------------
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_hdr = p_header.add_run("O'ZBEKISTON RESPUBLIKASI OLIY VA O'RTA MAXSUS TA'LIM VAZIRLIGI\nTA'LIM MUASSASASI")
    r_hdr.bold = True
    r_hdr.font.size = DocxPt(11)

    for _ in range(5):
        doc.add_paragraph()

    title_text = essay_data.get("title", "REFERAT").upper()
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t1 = p_t.add_run("REFERAT\n\n")
    r_t1.bold = True
    r_t1.font.size = DocxPt(22)

    r_t2 = p_t.add_run(f"MAVZU: {title_text}")
    r_t2.bold = True
    r_t2.font.size = DocxPt(14)

    for _ in range(6):
        doc.add_paragraph()

    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_inf = p_info.add_run(
        "Bajaruvchi: _____________________\n\n"
        "Guruh/Sinf: _____________________\n\n"
        "Tekshirdi: _____________________"
    )
    r_inf.font.size = DocxPt(12)

    for _ in range(5):
        doc.add_paragraph()

    p_ftr = doc.add_paragraph()
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ft = p_ftr.add_run("Toshkent - 2026")
    r_ft.font.size = DocxPt(11)

    doc.add_page_break()

    # ---------------------------------------------------------
    # PAGE 2: MUNDARIJA
    # ---------------------------------------------------------
    p_m_head = doc.add_paragraph()
    p_m_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_mh = p_m_head.add_run("MUNDARIJA")
    run_mh.bold = True
    run_mh.font.size = DocxPt(16)

    doc.add_paragraph()

    sections = essay_data.get("body_sections", [])
    p_m_list = doc.add_paragraph()
    p_m_list.paragraph_format.line_spacing = 1.5
    
    r_m = p_m_list.add_run("KIRISH .................................................................................................................... 3\n")
    r_m.font.size = DocxPt(12)

    page_num = 3
    for sec in sections:
        sec_title = sec.get("section_title", "Bob")
        r_m_sec = p_m_list.add_run(f"{sec_title} ................................................................................................... {page_num}\n")
        r_m_sec.font.size = DocxPt(12)
        page_num += 1

    r_m_c = p_m_list.add_run(f"XULOSA .................................................................................................................. {page_num}\n")
    r_m_c.font.size = DocxPt(12)
    
    r_m_r = p_m_list.add_run(f"FOYDALANILGAN ADABIYOTLAR RO'YXATI ..................................................... {page_num + 1}")
    r_m_r.font.size = DocxPt(12)

    doc.add_page_break()

    # ---------------------------------------------------------
    # PAGE 3: KIRISH
    # ---------------------------------------------------------
    p_ih = doc.add_paragraph()
    p_ih.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ih = p_ih.add_run("KIRISH")
    r_ih.bold = True
    r_ih.font.size = DocxPt(14)

    p_intro = doc.add_paragraph(essay_data.get("introduction", ""))
    p_intro.paragraph_format.line_spacing = 1.25
    p_intro.paragraph_format.space_after = DocxPt(12)

    doc.add_page_break()

    # ---------------------------------------------------------
    # PAGES 4-5: ASOSIY BOBLAR
    # ---------------------------------------------------------
    for sec in sections:
        p_sh = doc.add_paragraph()
        p_sh.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sh = p_sh.add_run(sec.get("section_title", "Bob").upper())
        r_sh.bold = True
        r_sh.font.size = DocxPt(14)

        p_s = doc.add_paragraph(sec.get("section_text", ""))
        p_s.paragraph_format.line_spacing = 1.25
        p_s.paragraph_format.space_after = DocxPt(12)

        doc.add_page_break()

    # ---------------------------------------------------------
    # PAGE 6: XULOSA VA ADABIYOTLAR
    # ---------------------------------------------------------
    p_ch = doc.add_paragraph()
    p_ch.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ch = p_ch.add_run("XULOSA")
    r_ch.bold = True
    r_ch.font.size = DocxPt(14)

    p_conc = doc.add_paragraph(essay_data.get("conclusion", ""))
    p_conc.paragraph_format.line_spacing = 1.25
    p_conc.paragraph_format.space_after = DocxPt(18)

    p_rh = doc.add_paragraph()
    p_rh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_rh = p_rh.add_run("FOYDALANILGAN ADABIYOTLAR RO'YXATI")
    r_rh.bold = True
    r_rh.font.size = DocxPt(14)

    refs = essay_data.get("references", ["1. Ilmiy darsliklar va manbalar."])
    for ref in refs:
        p_ref = doc.add_paragraph(ref)
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.paragraph_format.space_after = DocxPt(6)

    doc.save(output_filename)
    return output_filename

async def create_pptx_file(slides_data: list, output_filename: str) -> str:
    return await asyncio.to_thread(_build_pptx_sync, slides_data, output_filename)

async def create_docx_file(essay_data: dict, output_filename: str) -> str:
    return await asyncio.to_thread(_build_docx_sync, essay_data, output_filename)
