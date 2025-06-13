import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches
import zipfile
import io
import os
import tempfile
import re

st.set_page_config(page_title="Label Generator", layout="centered")
st.title("📦 Label Generator Web App")

uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
uploaded_logo = st.file_uploader("Upload Logo Image", type=["jpg", "jpeg", "png"])

if uploaded_excel and uploaded_logo:
    if st.button("Generate Labels"):
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded files
            excel_path = os.path.join(temp_dir, "order.xlsx")
            logo_path = os.path.join(temp_dir, "logo.jpg")
            with open(excel_path, "wb") as f:
                f.write(uploaded_excel.read())
            with open(logo_path, "wb") as f:
                f.write(uploaded_logo.read())

            # Read Excel file
            df = pd.read_excel(uploaded_excel)
            content_list = df.values.tolist()

            # Generate filename list (aa, ab, ...)
            alphabet = [chr(i) for i in range(97, 123)]
            l0 = [a + b for a in alphabet for b in alphabet]

            # Prepare output
            files_to_zip = []

            for i, row in enumerate(content_list):
                img = Image.open(uploaded_logo)
                draw = ImageDraw.Draw(img)

                try:
                    font = ImageFont.truetype("arial.ttf", 40)
                except:
                    font = ImageFont.load_default()

                draw.text((0, 370), f"{row[0]}  {row[1]}", fill=(0, 0, 0), font=font)
                draw.text((0, 440), str(row[2]), fill=(0, 0, 0), font=font)

                img_path = os.path.join(temp_dir, f"result{l0[i]}.jpg")
                img.save(img_path)

                # Word document
                doc = Document()
                doc.add_picture(img_path, width=Inches(3.6))
                safe_name = re.sub(r'[\\/*?:"<>|]', "_", str(row[0]))
                doc_path = os.path.join(temp_dir, f"{safe_name}.docx")
                doc.save(doc_path)

                files_to_zip.append(doc_path)

            # Create zip archive
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for file in files_to_zip:
                    zipf.write(file, arcname=os.path.basename(file))
            zip_buffer.seek(0)

            st.success("Done! Download your documents below.")
            st.download_button("📥 Download ZIP", zip_buffer, file_name="labels.zip")
