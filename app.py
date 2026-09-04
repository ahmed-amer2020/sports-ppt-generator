import streamlit as st
import docx
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

st.set_page_config(page_title="مُولد عروض التربية الرياضية", page_icon="🏆")
st.title("🏆 صانع العروض الأكاديمية - تربية رياضية")

api_key = st.text_input("أدخل مفتاح Gemini API الخاص بك:", type="password")
uploaded_file = st.file_uploader("اختر ملف البحث (Docx)", type=["docx"])

def read_docx(file):
    doc = docx.Document(file)
    return '\n'.join([para.text for para in doc.paragraphs])

def create_pptx(slides_data):
    prs = Presentation()
    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = slide_info.get("title", "شريحة بدون عنوان")
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 32, 67)
        
        txBox2 = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(9), Inches(5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for point in slide_info.get("content", []):
            p2 = tf2.add_paragraph()
            p2.text = f"• {point}"
            p2.font.size = Pt(20)
            p2.font.color.rgb = RGBColor(40, 40, 40)
            p2.space_after = Pt(14)
            
    binary_output = io.BytesIO()
    prs.save(binary_output)
    binary_output.seek(0)
    return binary_output

if uploaded_file is not None and api_key:
    if st.button("🚀 إنشاء عرض الباوربوينت"):
        with st.spinner("جاري معالجة البحث..."):
            try:
                text_content = read_docx(uploaded_file)
                genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-1.5-flash-latest')                
                prompt = f"""
                أنت خبير أكاديمي في كليات التربية الرياضية بمصر. قم بتلخيص هذا البحث واستخراج النقاط الأساسية لصناعة عرض PowerPoint لسيمينار أو مناقشة.
                أخرج النتيجة بنفس الصيغة التالية تماماً:
                
                الشريحة 1: عنوان البحث الباحث والمشرفين
                - [عنوان البحث]
                - إعداد الباحث: [الاسم إن وجد]
                
                الشريحة 2: مشكلة البحث وأهميته
                - [نقطة أساسية عن المشكلة]
                - [نقطة ثانية عن أهمية البحث]

                الشريحة 3: أهداف البحث وفروضه
                - [الهدف الرئيسي]
                - [الفروض الإحصائية]

                الشريحة 4: الإجراءات والمعالجات البدنية
                - منهج البحث والعينة
                - الاختبارات والقياسات المستخدمة

                الشريحة 5: أهم النتائج والتوصيات
                - [أهم نتيجة إحصائية]
                - [أهم توصية تطبيقية]

                نص البحث:
                {text_content[:4000]}
                """
                response = model.generate_content(prompt)
                ai_text = response.text
                
                slides = []
                current_slide = None
                for line in ai_text.split('\n'):
                    line = line.strip()
                    if line.startswith("الشريحة"):
                        if current_slide: slides.append(current_slide)
                        current_slide = {"title": line, "content": []}
                    elif line.startswith("-") and current_slide:
                        current_slide["content"].append(line.replace("-", "").strip())
                if current_slide: slides.append(current_slide)
                
                pptx_file = create_pptx(slides)
                st.success("تم إنشاء العرض بنجاح! 🎉")
                st.download_button("📥 تحميل ملف PowerPoint", pptx_file, "عرض_التربية_الرياضية.pptx")
            except Exception as e:
                st.error(f"حدث خطأ: {e}")
