import streamlit as st
import openai
import pandas as pd
from docx import Document
from docx.shared import RGBColor
import tempfile
import json
import re
import time
import os

# --- إعداد الواجهة ---
st.set_page_config(page_title="استخراج الكيانات من مستند Word", layout="centered")
st.title("📄 أداة استخراج الكيانات (NER)")

# --- إدخال البيانات ---
st.sidebar.header("إعدادات النموذج")
api_key = st.sidebar.text_input("🔑 أدخل مفتاح OpenAI API", type="password")
model = st.sidebar.selectbox("🤖 اختر النموذج", ["gpt-3.5-turbo", "gpt-4"])

uploaded_file = st.file_uploader("📤 ارفع ملف Word (بصيغة .docx)", type=["docx"])

# --- إعداد الأنماط ---
style_map = {
    "الأماكن": {"style": "أماكن", "color": RGBColor(255, 0, 0)},
    "الأعلام": {"style": "أعلام", "color": RGBColor(0, 0, 255)},
    "الفرق": {"style": "فرق", "color": RGBColor(0, 128, 0)},
    "الكتب": {"style": "كتب", "color": RGBColor(128, 0, 128)},
    "الشواهد": {"style": "شواهد", "color": RGBColor(255, 165, 0)},
}

# --- دالة استخراج الكيانات ---
def extract_entities(client, model, paragraph_text):
    prompt = f"""
أنت خبير لغوي مكلف بتحليل النص التالي واستخراج الكيانات التالية بدقة:
1. الأماكن
2. الأعلام
3. الفرق
4. الكتب
5. الشواهد (آيات أو أحاديث)

النص:
{paragraph_text}

الرجاء الرد بصيغة JSON فقط كما يلي:
{{
  "الأماكن": [],
  "الأعلام": [],
  "الفرق": [],
  "الكتب": [],
  "الشواهد": []
}}
"""
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
    )
    raw_content = response.choices[0].message.content
    cleaned = raw_content.replace("،", ",").replace("“", '"').replace("”", '"')
    match = re.search(r'\{[\s\S]*\}', cleaned)
    if match:
        return json.loads(match.group())
    else:
        return {}

# --- تنسيق الكيانات ---
def rewrite_paragraph_with_styles(doc, text, entities):
    new_paragraph = doc.add_paragraph()
    new_runs = []
    pointer = 0
    matches = []

    for entity_type, values in entities.items():
        for val in values:
            start = text.find(val, pointer)
            if start != -1:
                matches.append((start, start + len(val), val, entity_type))

    matches.sort()
    for start, end, val, etype in matches:
        if start > pointer:
            new_runs.append(("plain", text[pointer:start]))
        new_runs.append((etype, val))
        pointer = end
    if pointer < len(text):
        new_runs.append(("plain", text[pointer:]))

    for style_type, content in new_runs:
        run = new_paragraph.add_run(content)
        if style_type != "plain":
            run.style = style_map[style_type]["style"]
            run.font.color.rgb = style_map[style_type]["color"]

# --- المعالجة ---
if uploaded_file and api_key:
    with st.spinner("🔍 جاري تحليل الملف..."):
        client = openai.OpenAI(api_key=api_key)
        input_doc = Document(uploaded_file)
        template_path = "template.docx"

        if not os.path.exists(template_path):
            st.error("❌ ملف القالب template.docx غير موجود. الرجاء إضافته إلى مجلد المشروع.")
        else:
            new_doc = Document(template_path)
            all_entities = []

            for para in input_doc.paragraphs:
                text = para.text.strip()
                if not text:
                    continue

                entities = extract_entities(client, model, text)
                if entities:
                    rewrite_paragraph_with_styles(new_doc, text, entities)
                    for etype, vals in entities.items():
                        for v in vals:
                            all_entities.append({"الكيان": v, "النوع": etype, "الفقرة": text})

                time.sleep(1.5)

            # --- حفظ الملفات ---
            word_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            new_doc.save(word_file.name)

            df = pd.DataFrame(all_entities)
            excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df.to_excel(excel_file.name, index=False)

            st.success("✅ تم التحليل بنجاح")
            st.download_button("📥 تحميل ملف Word الناتج", data=open(word_file.name, "rb"), file_name="نتيجة.docx")
            st.download_button("📊 تحميل ملف Excel للكيانات", data=open(excel_file.name, "rb"), file_name="الكيانات.xlsx")

elif uploaded_file and not api_key:
    st.warning("⚠️ الرجاء إدخال مفتاح API للمتابعة.")
