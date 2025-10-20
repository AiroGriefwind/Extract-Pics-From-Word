import streamlit as st
import zipfile
import io
import re

def sanitize_filename(s: str) -> str:
    # Keep CJK, numbers, letters, dash, and space, then replace spaces with dash
    s = re.sub(r'[^\u4e00-\u9fff\w\- ]', '', s)
    return s.strip().replace(' ', '-')

st.set_page_config(page_title="DOCX Image Extractor", page_icon="📄")
st.title("DOCX Image Extractor")

uploaded = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded is not None:
    # Read file bytes into memory
    file_bytes = uploaded.read()
    doc_title = sanitize_filename(uploaded.name.rsplit('.', 1)[0]) or "document"

    # Open the .docx as a zip and find images in word/media
    with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as docx_zip:
        media_files = sorted([f for f in docx_zip.namelist() if f.startswith('word/media/')])

        if not media_files:
            st.warning("No embedded images were found in this document.")
        else:
            # Build an in-memory zip of extracted images with new names
            out_buf = io.BytesIO()
            with zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED) as out_zip:
                for idx, image_path in enumerate(media_files, 1):
                    ext = ('.' + image_path.split('.')[-1].lower()) if '.' in image_path else ''
                    new_name = f"{doc_title}-{idx}{ext}"
                    image_data = docx_zip.read(image_path)
                    out_zip.writestr(new_name, image_data)

            out_buf.seek(0)
            download_name = f"{doc_title}-images.zip"

            st.success(f"Extracted {len(media_files)} images.")
            st.download_button(
                "Download ZIP",
                data=out_buf.getvalue(),
                file_name=download_name,
                mime="application/zip",
            )
