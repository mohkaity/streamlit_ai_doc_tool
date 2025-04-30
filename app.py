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
st.set_page_config(page_title="تحليل المستند بالذكاء الاصطناعي", layout="centered")
st.title("📄 أداة تحليل الفقرات واستخراج الكيانات")

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

# --- دوال GPT ---
def generate_title(client, model, paragraph_text):
    prompt = f"""
اقرأ النص التالي واستخرج منه عنوانًا موجزًا يدل على فكرته الأساسية. لا تشرح، فقط أعطني العنوان نفسه دون علامات تنصيص.
الرجاء الرد فقط بالعنوان دون أي مقدمة أو تعليق.
النص:
{paragraph_text}
"""
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    return response.choices[0].message.content.strip().strip('"“”')

def extract_entities(client, model, paragraph_text):
    prompt = f"""
    أنت خبير لغوي مكلف بتحليل النص التالي واستخراج الكيانات التالية بدقة:

           1. الأماكن (المدن، الدول، المناطق)
              - ❌ لا تدرج أسماء منسوبة (مثل: البغداديين)
              - ✅ أدرج فقط الاسم الأصلي مثل: "بغداد"، "مكة"
              بمعنى أنك لو وجدت كلمة تعني نسبة شخص لمنطقة فلا تعتبرها مكاناً مثلا البغدادي لا يعني بغداد
              فقط ضع اعتبار للمكان المجرد وأعتبره من الأماكن سواء كان مباشراً مثل يمن أو معرفاً مثل اليمن لكن لو كان نسبة مثل اليمني فلا تضفه

           2. الأعلام (أسماء أشخاص حقيقيين فقط)
              - ❌ لا تدرج ألقاب جماعية أو صفات نسبية مثل: "الشافعية"، "الحنابلة"
              - ✅ فقط الأسماء الصريحة مثل: "ابن تيمية"، "أبو حامد الغزالي"

           3. الفرق والمذاهب (الكيانات الدينية والفكرية المعروفة)
              - ✅ مثل: "المرجئة"، "الأشاعرة"، "المعتزلة"

           4. أسماء الكتب (عناوين كتب علمية أو شرعية)
              - ✅ مثل: "صحيح البخاري"، "الإحياء"، "الرد على المنطقيين"

           5. الشواهد (آيات أو أحاديث فقط)
              - ✅ كل ما يبدأ بـ "قال الله تعالى" أو "قال رسول الله"

           ❗ المطلوب: استخراج الكيانات بدقة وتجاهل أي كلمات غير مطابقة للمعايير السابقة.

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
def rewrite_paragraph_with_styles(paragraph, entities):
    full_text = paragraph.text
    new_runs = []
    pointer = 0
    matches = []

    for entity_type, values in entities.items():
        for val in values:
            start = full_text.find(val, pointer)
            if start != -1:
                matches.append((start, start + len(val), val, entity_type))

    matches.sort()
    for start, end, val, etype in matches:
        if start > pointer:
            new_runs.append(("plain", full_text[pointer:start]))
        new_runs.append((etype, val))
        pointer = end
    if pointer < len(full_text):
        new_runs.append(("plain", full_text[pointer:]))

    for run in paragraph.runs:
        run.text = ""

    for style_type, content in new_runs:
        run = paragraph.add_run(content)
        if style_type != "plain":
            run.style = style_map[style_type]["style"]
            run.font.color.rgb = style_map[style_type]["color"]

# --- المعالجة ---
if uploaded_file and api_key:
    with st.spinner("🔍 جاري تحليل الملف..."):
        client = openai.OpenAI(api_key=api_key)
        doc = Document(uploaded_file)
        new_doc = Document()
        all_entities = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue

            title = generate_title(client, model, text)
            title_para = new_doc.add_paragraph()
            title_run = title_para.add_run(title)
            title_run.bold = True

            body_para = new_doc.add_paragraph(text)
            entities = extract_entities(client, model, text)
            all_entities.append({"الكيان": f"[عنوان] {title}", "النوع": "عنوان", "الفقرة": text})

            if entities:
                rewrite_paragraph_with_styles(body_para, entities)
                for etype, vals in entities.items():
                    for v in vals:
                        all_entities.append({"الكيان": v, "النوع": etype, "الفقرة": text})

            time.sleep(1.5)

        # --- حفظ الملفات المؤقتة ---
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
