import streamlit as st
import docx
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import io

# 1. إعدادات الصفحة
st.set_page_config(page_title="صانع العروض التوضيحية الاحترافي - تربية رياضية", page_icon="🏆", layout="wide")

st.title("🏆 صانع العروض البصرية والأكاديمية - نمط NotebookLM")
st.write("حمل ملف البحث الخاص بك للحصول على عرض بصري بهيكلية الإنفوجرافيك والبطاقات التوضيحية الأنيقة.")

# 2. البيانات والمدخلات
api_key = st.text_input("🔑 أدخل مفتاح Gemini API الخاص بك:", type="password")

st.markdown("---")
st.subheader("📋 البيانات التوجيهية للعرض")

col1, col2 = st.columns(2)
with col1:
    research_type = st.selectbox(
        "🎯 نوع العرض المطلوب:",
        ["مناقشة رسالة (ماجستير / دكتوراه)", "سيمنار تسجيل خطة بحث (Proposal)", "عرض تدريبي / ورقة عمل"]
    )
    specialty = st.selectbox(
        "⚽ التخصص الرياضي الدقيق:",
        ["التدريب الرياضي وعلوم الحركة", "الإدارة الرياضية والترويح", "المناهج وطرق التدريس", "علوم الصحة والتربية البدنية", "علم النفس الرياضي"]
    )
    slides_count = st.select_slider("📊 عدد الشرائح المطلوب:", options=[5, 8, 10, 12], value=8)

with col2:
    research_title = st.text_input("📌 عنوان البحث / الموضوع الرئيسي:", placeholder="مثل: مبدأ التدريب بناءً على الاستشفاء")
    researcher_name = st.text_input("👨‍🎓 إعداد (الباحث / الكابتن):", placeholder="مثل: كابتن / عامر")
    supervisors = st.text_input("👨‍🏫 تحت إشراف:", placeholder="مثل: دكتور محمد إبراهيم جعفر")
    logo_file = st.file_uploader("🖼️ رفع شعار/لوجو العرض (اختياري - PNG/JPG)", type=["png", "jpg", "jpeg"])

uploaded_file = st.file_uploader("📂 اختر ملف البحث (Docx)", type=["docx"])

def read_docx(file):
    doc = docx.Document(file)
    fullText = []
    for para in doc.paragraphs:
        if para.text.strip():
            fullText.append(para.text.strip())
    return '\n'.join(fullText)

def create_pptx(slides_data, meta_info, logo_bytes=None):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    BG_COLOR = RGBColor(240, 242, 245)
    HEADER_TEXT_COLOR = RGBColor(30, 41, 59)
    ORANGE_ACCENT = RGBColor(234, 88, 12)
    CARD_BG = RGBColor(255, 255, 255)
    BORDER_COLOR = RGBColor(203, 213, 225)
    
    for idx, slide_info in enumerate(slides_data):
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # خلفية الشريحة العامة
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = BG_COLOR
        bg_shape.line.fill.background()
        
        # إضافة الشعار العلوي إن وجد
        if logo_bytes:
            try:
                logo_stream = io.BytesIO(logo_bytes)
                slide.shapes.add_picture(logo_stream, Inches(11.8), Inches(0.3), Inches(1.1), Inches(1.1))
            except Exception:
                pass
        
        # ---------------- 1. شريحة الغلاف الرئيسية ----------------
        if idx == 0:
            bg_shape.fill.fore_color.rgb = RGBColor(15, 23, 42)
            
            txBox = slide.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(2.5))
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = meta_info['title'] if meta_info['title'] else slide_info.get("title", "العنوان الرئيسي")
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.alignment = PP_ALIGN.RIGHT
            
            accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(4.2), Inches(11.333), Inches(0.08))
            accent_bar.fill.solid()
            accent_bar.fill.fore_color.rgb = ORANGE_ACCENT
            accent_bar.line.fill.background()
            
            txBox2 = slide.shapes.add_textbox(Inches(1.0), Inches(4.8), Inches(11.333), Inches(2.0))
            tf2 = txBox2.text_frame
            
            p_res = tf2.paragraphs[0]
            p_res.text = f"إعداد: {meta_info['researcher']}" if meta_info['researcher'] else ""
            p_res.font.size = Pt(22)
            p_res.font.bold = True
            p_res.font.color.rgb = ORANGE_ACCENT
            p_res.alignment = PP_ALIGN.RIGHT
            
            if meta_info['supervisors']:
                p_sup = tf2.add_paragraph()
                p_sup.text = f"تحت إشراف: {meta_info['supervisors']}"
                p_sup.font.size = Pt(20)
                p_sup.font.bold = True
                p_sup.font.color.rgb = RGBColor(226, 232, 240)
                p_sup.alignment = PP_ALIGN.RIGHT
                p_sup.space_before = Pt(10)

        # ---------------- 2. باقي الشرائح المصممة كبطاقات إنفوجرافيك ----------------
        else:
            slide_num_str = f"0{idx+1}" if idx < 9 else f"{idx+1}"
            tx_head = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(10.5), Inches(1.0))
            tf_head = tx_head.text_frame
            tf_head.word_wrap = True
            p_head = tf_head.paragraphs[0]
            p_head.text = f"{slide_num_str} | {slide_info.get('title', '')}"
            p_head.font.size = Pt(28)
            p_head.font.bold = True
            p_head.font.color.rgb = HEADER_TEXT_COLOR
            p_head.alignment = PP_ALIGN.RIGHT
            
            content_points = slide_info.get("content", [])
            num_points = len(content_points)
            
            # تقسيم المحتوى إلى بطاقات جانبية أو أعمدة مقسمة تلقائياً
            if num_points <= 3:
                card_width = Inches(11.733)
                card_height = Inches(1.3)
                top_pos = Inches(1.6)
                
                for pt in content_points:
                    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), top_pos, card_width, card_height)
                    card.fill.solid()
                    card.fill.fore_color.rgb = CARD_BG
                    card.line.color.rgb = BORDER_COLOR
                    
                    tx = slide.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.15), Inches(11.333), card_height)
                    tf_card = tx.text_frame
                    tf_card.word_wrap = True
                    p_card = tf_card.paragraphs[0]
                    p_card.text = f"• {pt}"
                    p_card.font.size = Pt(18)
                    p_card.font.color.rgb = HEADER_TEXT_COLOR
                    p_card.alignment = PP_ALIGN.RIGHT
                    
                    top_pos += Inches(1.6)
            else:
                col_width = Inches(5.6)
                col_gap = Inches(0.533)
                
                for i, pt in enumerate(content_points[:4]):
                    col_idx = i % 2
                    row_idx = i // 2
                    
                    left_pos = Inches(0.8) + col_idx * (col_width + col_gap) if col_idx == 1 else Inches(0.8)
                    top_pos = Inches(1.8) + row_idx * Inches(2.5)
                    
                    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_pos, top_pos, col_width, Inches(2.2))
                    card.fill.solid()
                    card.fill.fore_color.rgb = CARD_BG
                    card.line.color.rgb = BORDER_COLOR
                    
                    tx = slide.shapes.add_textbox(left_pos + Inches(0.2), top_pos + Inches(0.2), col_width - Inches(0.4), Inches(1.8))
                    tf_card = tx.text_frame
                    tf_card.word_wrap = True
                    p_card = tf_card.paragraphs[0]
                    p_card.text = f"• {pt}"
                    p_card.font.size = Pt(17)
                    p_card.font.color.rgb = HEADER_TEXT_COLOR
                    p_card.alignment = PP_ALIGN.RIGHT
            
    binary_output = io.BytesIO()
    prs.save(binary_output)
    binary_output.seek(0)
    return binary_output

if uploaded_file is not None and api_key:
    if st.button("🚀 إنشاء وتصميم العرض البصري"):
        with st.spinner("جاري صياغة المحتوى وبناء شرائح الإنفوجرافيك..."):
            try:
                text_content = read_docx(uploaded_file)
                genai.configure(api_key=api_key)
                
                available_models = ['models/gemini-2.5-flash', 'models/gemini-2.0-flash', 'models/gemini-1.5-flash']
                
                prompt = f"""
                أنت مصمم إنفوجرافيك خبير ومحاضر في كليات التربية الرياضية تعمل بأسلوب NotebookLM.
                قم بتقسيم وتلخيص هذا البحث إلى عرض PowerPoint بصري وهيكلي حديث لـ ({research_type}).
                
                المعطيات:
                - التخصص: {specialty}
                - عدد الشرائح: {slides_count}
                
                الشروط:
                1. صغ النقاط على شكل عبارات إنفوجرافية مكثفة وقصيرة (شديدة الوضوح والتركيز).
                2. ابدأ كل شريحة بـ "شريحة:" يتبعها العنوان الهيكلي المباشر (مثل: قانون التكيف التعويضي، الفاصل البيولوجي، إلخ).
                3. تجنب الفقرات الطويلة واجعل كل نقطة تعبر عن مفهوم محدد قابل للعرض في بطاقة.

                الصيغة المطلوبة:
                شريحة: [عنوان الشريحة]
                - [مفهوم أو عنصر مكثف]
                - [مفهوم أو عنصر آخر]

                نص البحث:
                {text_content[:5000]}
                """
                
                response = None
                for model_name in available_models:
                    try:
                        model = genai.GenerativeModel(model_name)
                        response = model.generate_content(prompt)
                        if response and response.text:
                            break
                    except Exception:
                        continue
                
                if response:
                    ai_text = response.text
                    slides = []
                    current_slide = None
                    
                    for line in ai_text.split('\n'):
                        line = line.strip()
                        if line.startswith("شريحة:") or line.startswith("الشريحة"):
                            if current_slide and current_slide["content"]:
                                slides.append(current_slide)
                            clean_title = line.split(":", 1)[-1].strip() if ":" in line else line
                            current_slide = {"title": clean_title, "content": []}
                        elif line.startswith("-") and current_slide:
                            point_text = line.lstrip("-").strip()
                            if point_text:
                                current_slide["content"].append(point_text)
                    
                    if current_slide and current_slide["content"]:
                        slides.append(current_slide)
                    
                    meta_info = {
                        'title': research_title,
                        'researcher': researcher_name,
                        'supervisors': supervisors,
                        'specialty': specialty
                    }
                    
                    logo_bytes = logo_file.getvalue() if logo_file else None
                    
                    st.session_state['generated_slides'] = slides
                    st.session_state['meta_info'] = meta_info
                    st.session_state['pptx_bytes'] = create_pptx(slides, meta_info, logo_bytes)
                    
                    st.success("🎉 تم إنشاء العرض التوضيحي بنجاح!")

            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")

# 3. المعاينة والتحميل
if 'generated_slides' in st.session_state:
    st.markdown("---")
    st.subheader("👁️ معاينة البطاقات والشرائح قبل التحميل")
    
    slides = st.session_state['generated_slides']
    meta = st.session_state['meta_info']
    
    slide_idx = st.slider("اختر الشريحة للمعاينة:", 1, len(slides), 1) - 1
    selected_slide = slides[slide_idx]
    
    with st.container():
        st.markdown(
            f"""
            <div style="border: 2px solid #CBD5E1; border-radius: 12px; padding: 25px; background-color: #F0F2F5; direction: rtl; text-align: right;">
                <div style="color: #1E293B; font-size: 24px; font-weight: bold; margin-bottom: 15px;">
                    0{slide_idx+1} | {selected_slide['title']}
                </div>
                <hr style="border: 1px solid #EA580C; margin-bottom: 20px;">
                <div style="font-size: 18px; color: #333; line-height: 1.8;">
            """,
            unsafe_allow_html=True
        )
        
        for p in selected_slide['content']:
            st.markdown(f"<div style='background: white; padding: 12px; border-radius: 8px; margin-bottom: 10px; border-right: 5px solid #EA580C;'>• {p}</div>", unsafe_allow_html=True)
            
        st.markdown("</div></div>", unsafe_allow_html=True)
        
    st.write("")
    st.download_button(
        label="📥 تحميل ملف PowerPoint الإنفوجرافيكي (.pptx)",
        data=st.session_state['pptx_bytes'],
        file_name="عرض_إنفوجرافيك_التربية_الرياضية.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
