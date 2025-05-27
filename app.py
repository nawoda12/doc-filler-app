import streamlit as st
from docx import Document
from docx.shared import RGBColor, Inches
import re
from datetime import datetime

def fill_template(template_path, replacements, logo_path=None):
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
                if "THE MASTER AGREEMENT AND" not in run.text:
                    run.font.color.rgb = RGBColor(0, 0, 0)

    # Handle tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text
                new_cell_text = cell_text

                # Replace placeholders in tables
                for placeholder, value in replacements.items():
                    if placeholder in new_cell_text:
                        new_cell_text = new_cell_text.replace(placeholder, value)

                # Remove unreplaced placeholders and brackets in tables
                new_cell_text = re.sub(r'<[^>]*>', '', new_cell_text)
                new_cell_text = re.sub(r'\[[^\]]*\]', '', new_cell_text)

                cell.text = new_cell_text
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0)

    # Insert logo if provided
    if logo_path:
        for para in doc.paragraphs:
            if '[LOGO]' in para.text:
                para.clear()
                run = para.add_run()
                run.add_picture(logo_path, width=Inches(2))

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
    replacements["_______________________________________________"] = replacements.get("Customer", "XYZ Co.")
    replacements["______________________________________________ _____________"] = replacements.get("Customer Address", "330, Flatbush Avenue, Brooklyn, New York 11238, USA")
    replacements["__________________"] = replacements.get("Contract Execution Date", "CONTRACT EXECUTION DATE")
    replacements["[year]"] = str(datetime.now().year)
    return replacements

st.title("📄 Statement of Work Auto-Filler")

uploaded_file = st.file_uploader("Upload your Word template (.docx)", type="docx")
logo_file = st.file_uploader("Upload your logo (.png, .jpg)", type=["png", "jpg"])
email_content = st.text_area("Paste the client email content here")

if uploaded_file and email_content:
    replacements = extract_replacements(email_content)
    logo_path = None
    if logo_file:
        logo_path = "uploaded_logo." + logo_file.name.split('.')[-1]
        with open(logo_path, "wb") as f:
            f.write(logo_file.getbuffer())
    filled_doc = fill_template(uploaded_file, replacements, logo_path)
    filled_doc_path = "Filled_Statement_of_Work.docx"
    filled_doc.save(filled_doc_path)

    with open(filled_doc_path, "rb") as file:
        st.download_button(
            label="📥 Download Filled Document",
            data=file,
            file_name="Filled_Statement_of_Work.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
