import streamlit as st
from docx import Document
from docx.shared import RGBColor
import re

def fill_template(template_path, replacements):
    doc = Document(template_path)
    for para in doc.paragraphs:
        for run in para.runs:
            for placeholder, value in replacements.items():
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
                    if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                        run.font.color.rgb = RGBColor(255, 0, 0)
    return doc

def extract_replacements(email_content):
    replacements = {}
    lines = email_content.split('\n')
    for line in lines:
        match = re.match(r'(.+?):\s*(.+)', line)
        if match:
            key = match.group(1).strip()
            value = match.group(2).strip()
            replacements[f"<{key}>"] = value
    return replacements

st.title("📄 Document Auto-Filler")

uploaded_file = st.file_uploader("Upload your Word template (.docx)", type="docx")
email_content = st.text_area("Paste the client email content here")

if uploaded_file and email_content:
    replacements = extract_replacements(email_content)
    filled_doc = fill_template(uploaded_file, replacements)
    filled_doc_path = "Filled_Statement_of_Work.docx"
    filled_doc.save(filled_doc_path)

    with open(filled_doc_path, "rb") as file:
        st.download_button(
            label="📥 Download Filled Document",
            data=file,
            file_name="Filled_Statement_of_Work.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
