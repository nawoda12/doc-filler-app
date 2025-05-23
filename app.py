import streamlit as st
from docx import Document
from docx.shared import RGBColor
import re

def fill_template(template_path, replacements):
    doc = Document(template_path)
    for para in doc.paragraphs:
        original_text = para.text
        new_text = original_text

        # Replace placeholders
        for placeholder, value in replacements.items():
            if placeholder in new_text:
                new_text = new_text.replace(placeholder, value)

        # Remove unreplaced placeholders and brackets
        new_text = re.sub(r'<[^>]*>', '', new_text)
        new_text = re.sub(r'\[[^\]]*\]', '', new_text)

        # Replace 'xx Hours' if 'Total Estimated Hours' is present
        if "Total Estimated Hours" in new_text and "xx Hours" in new_text and "Hours" in replacements:
            new_text = new_text.replace("xx Hours", replacements["Hours"])

        # Remove lines with empty values
        if any(key in original_text for key in replacements.keys()) or original_text.strip() != "":
            para.text = new_text
            for run in para.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)

    return doc

def extract_replacements(email_content):
    replacements = {
        "<Name>": "Nawoda Sathsara"
    }
    lines = email_content.split('\n')
    for line in lines:
        match = re.match(r'(.+?):\s*(.+)', line)
        if match:
            key = match.group(1).strip()
            value = match.group(2).strip()
            replacements[f"<{key}>"] = value
            replacements[f"[{key}]"] = value
            replacements[key] = value
    # Add common blanks
    replacements["_____________"] = replacements.get("Date", "DATE")
    replacements["_______________________________________________"] = replacements.get("Customer", "CUSTOMER NAME")
    replacements["______________________________________________ _____________"] = replacements.get("Customer Address", "CUSTOMER ADDRESS")
    replacements["__________________"] = replacements.get("Contract Execution Date", "CONTRACT EXECUTION DATE")
    return replacements

st.title("📄 Statement of Work Auto-Filler")

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
