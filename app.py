import streamlit as st
import os
import re
import docx
from weasyprint import HTML
from google import genai

# ضبط إعدادات الصفحة في Streamlit
st.set_page_config(
    page_title="صانع العروض التقديمية | NotebookLM Style",
    page_icon="🎨",
    layout="wide"
)

st.title("🎨 صانع العروض التقديمية الاحترافية (NotebookLM Style)")
st.caption("رفع ملف البحث (Word/Docx) وسيقوم الذكاء الاصطناعي بتلخيصه وتوليد عرض بصري PDF بنفس أسلوب البطاقات التفاعلية.")

# إدخال مفتاح API الخاص بـ Gemini
api_key = st.sidebar.text_input("مفتاح Gemini API Key:", type="password")

# دالة لقراءة النص من ملف Word
def read_docx(file):
    doc = docx.Document(file)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return "\n".join(full_text)

# دالة الجلب التلقائي والديناميكي لأفضل موديل شغال ومتاح في حسابك
def get_working_model_name(client):
    try:
        models = list(client.models.list())
        flash_models = []
        general_models = []
        
        for m in models:
            # استخراج المعرف أو الاسم
            name = getattr(m, 'name', '') or getattr(m, 'id', '')
            # التحقق من دعم التوليد
            methods = getattr(m, 'supported_generation_methods', []) or getattr(m, 'supported_actions', [])
            
            # إذا كان الموديل يدعم generateContent
            if not methods or 'generateContent' in methods:
                clean_name = name.replace("models/", "")
                if 'flash' in clean_name.lower():
                    flash_models.append(clean_name)
                elif 'gemini' in clean_name.lower():
                    general_models.append(clean_name)

        # نفضل الموديلات السريعة (flash)، ثم أي موديل gemini آخر
        if flash_models:
            return flash_models[0]
        if general_models:
            return general_models[0]
            
    except Exception as e:
        st.warning(f"⚠️ تعذر جلب قائمة الموديلات تلقائياً: {e}، سنحاول استخدام الموديل الافتراضي.")
    
    # خيار fallback احتياطي
    return "gemini-1.5-flash"

# دالة استخراج كود HTML فقط وتنظيفه من أي نصوص زائدة
def extract_clean_html(text):
    match = re.search(r'<!DOCTYPE html>.*</html>', text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(0)
    clean_text = re.sub(r'^```html', '', text, flags=re.MULTILINE)
    clean_text = re.sub(r'^```', '', clean_text, flags=re.MULTILINE)
    return clean_text.strip()

# دالة توليد كود HTML/CSS الشرائح
def generate_presentation_html(text_content, api_key):
    client = genai.Client(api_key=api_key)
    
    # 🔍 التحديد التلقائي والديناميكي للموديل المتاح في حسابك
    active_model = get_working_model_name(client)
    st.info(f"🤖 الموديل النشط والمستخدم تلقائياً: `{active_model}`")
    
    prompt = f"""
    أنت مصمم عروض تقديمية أكاديمية وخبير بالهيكلة البصرية بأسلوب NotebookLM.
    قم بتحليل النص البحثي التالي واستخرج منه المحاور الرئيسية، ثم قم بتوليد عرض تقديمي بصري محول إلى كود HTML/CSS متكامل للطباعة بصيغة PDF A4 Landscape.

    النص البحثي:
    \"\"\"{text_content[:8000]}\"\"\"

    الشروط القاسية في كود HTML المطلوب:
    1. اتجاه الصفحة RTL واللغة العربية dir="rtl".
    2. صمم غلاف متدرج داكن (Dark Gradient Cover) بنفس تصميم الشريحة الأولى.
    3. قسم باقي الشرائح (على الأقل 4 شرائح رئيسية) إلى بطاقات أنيقة (Cards) ذات حواف جانبية بارزة ومربعات إحصائية/أرقام (Metric Boxes) ومقارنات بأعمدة.
    4. استخدم نظام الألوان الداكنة والحيوية (الكحلي الداكن #0f172a، البرتقالي #ea580c، الأزرق #0284c7، الأخضر #10b981).
    5. قم بإنهاء الكود بالكامل بـ </html> ولا تضف أي شرح خارجي قبل أو بعد الكود.

    استخدم الهيكل التالي كقاعدة صلبة لتنفيذ الشرائح:
    <!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4 landscape; margin: 0; background-color: #0f172a; }}
            *, *::before, *::after {{ box-sizing: border-box; }}
            body {{ margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, sans-serif; background-color: #0f172a; color: #334155; }}
            .slide {{ width: 297mm; height: 210mm; page-break-after: always; position: relative; background-color: #f8fafc; overflow: hidden; padding: 18mm 22mm; }}
            
            /* Slide Cover */
            .slide-cover {{ background: linear-gradient(135deg, #0f172a 0%, #1e293b 60%, #0f172a 100%); display: table; width: 297mm; height: 210mm; padding: 25mm 22mm; }}
            .cover-content {{ display: table-cell; vertical-align: middle; }}
            .cover-tag {{ display: inline-block; background: linear-gradient(90deg, #ea580c, #f97316); color: white; padding: 8px 20px; border-radius: 20px; font-size: 13pt; font-weight: bold; margin-bottom: 25px; }}
            .cover-title {{ color: #ffffff; font-size: 30pt; font-weight: 800; line-height: 1.35; margin-bottom: 18px; }}
            .cover-subtitle {{ color: #94a3b8; font-size: 16pt; margin-bottom: 35px; }}
            
            /* Inner Slide Elements */
            .header {{ border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 22px; }}
            .slide-num {{ color: #ea580c; font-size: 18pt; font-weight: 800; margin-left: 10px; }}
            .slide-title {{ color: #0f172a; font-size: 22pt; font-weight: 700; display: inline; }}
            
            .card-item {{ background: #ffffff; border-radius: 12px; padding: 16px 22px; margin-bottom: 14px; border-right: 6px solid #ea580c; box-shadow: 0 4px 12px rgba(0,0,0,0.03); border-top: 1px solid #f1f5f9; border-left: 1px solid #f1f5f9; border-bottom: 1px solid #f1f5f9; }}
            .card-header {{ font-size: 15pt; font-weight: 700; color: #0f172a; margin-bottom: 6px; }}
            .card-body {{ font-size: 12pt; color: #475569; line-height: 1.6; }}
            
            .cols-table {{ width: 100%; border-collapse: separate; border-spacing: 16px 0; }}
            .col-card {{ width: 50%; vertical-align: top; background: #ffffff; border-radius: 14px; padding: 20px; border-top: 5px solid #0284c7; box-shadow: 0 4px 14px rgba(0,0,0,0.04); }}
            .col-card.orange {{ border-top-color: #ea580c; }}
            
            .metric-box {{ background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 10px; padding: 12px; text-align: center; margin-top: 14px; }}
            .metric-value {{ font-size: 20pt; font-weight: 800; color: #0369a1; }}
            .metric-desc {{ font-size: 11pt; color: #075985; margin-top: 2px; }}
            
            .footer {{ position: absolute; bottom: 10mm; left: 22mm; right: 22mm; font-size: 10pt; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 8px; display: table; width: calc(100% - 44mm); }}
            .footer-left {{ display: table-cell; text-align: left; }}
            .footer-right {{ display: table-cell; text-align: right; }}
        </style>
    </head>
    <body>
        <!-- مولد الشرائح هنا -->
    </body>
    </html>
    """
    
    response = client.models.generate_content(
        model=active_model,
        contents=prompt
    )
    
    return extract_clean_html(response.text)

# واجهة رفع الملفات
uploaded_file = st.file_uploader("قم برفع ملف البحث (Docx):", type=["docx"])

if uploaded_file and api_key:
    if st.button("🚀 ابدأ توليد العرض التقديمي"):
        try:
            with st.spinner("1️⃣ جاري قراءة البحث وتحليله..."):
                text_content = read_docx(uploaded_file)
                st.info(f"تم قراءة {len(text_content)} حرف من الملف.")
                
            with st.spinner("2️⃣ جاري تصميم كود HTML وتحويله إلى PDF بصري..."):
                html_code = generate_presentation_html(text_content, api_key)
                
                # تحويل HTML إلى PDF
                pdf_filename = "presentation_notebooklm.pdf"
                HTML(string=html_code).write_pdf(pdf_filename)
                
                st.success("✨ تم توليد العرض التقديمي بنجاح!")
                
                # زر تحميل الملف الناتج
                with open(pdf_filename, "rb") as file:
                    st.download_button(
                        label="📥 تحميل العرض التقديمي (PDF)",
                        data=file,
                        file_name="عرض_بحثي_احترافي.pdf",
                        mime="application/pdf"
                    )
        except Exception as e:
            st.error(f"❌ حدث خطأ أثناء التوليد: {str(e)}")
elif not api_key:
    st.info("👈 يرجى إدخال مفتاح Gemini API Key في القائمة الجانبية للبدء.")
