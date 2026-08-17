"""
app.py
------
The web "platform". Run this with:  streamlit run app.py
It gives you a page in your browser with an "Upload Excel" button plus
fields for the executive's LinkedIn ID, name, country, and photo.
Upload everything, click Generate, and download the finished .pptx.
"""

import streamlit as st
from ppt_generator import generate_pptx

st.set_page_config(page_title="LinkedIn Analysis PPT Generator", page_icon="📊")

st.title("📊 LinkedIn Analysis PPT Generator")
st.write(
    "Upload the Excel file (same columns as `LA.xlsx`: PostType, TimeAgo, "
    "Content, PostLink, Source, Headline), fill in the executive's details "
    "below, and this tool will fill the `Sample.pptx` template with "
    "everything and give you a finished PowerPoint to download."
)

TEMPLATE_PATH = "Sample.pptx"  # must sit next to app.py in the repo

st.subheader("1. Executive details")
col1, col2 = st.columns(2)
with col1:
    exec_name = st.text_input("Executive name", placeholder="e.g. Jane Doe")
    linkedin_id = st.text_input(
        "LinkedIn ID",
        placeholder="e.g. jane-doe-exec-12345/",
        help="The part of the LinkedIn profile URL after linkedin.com/in/",
    )
with col2:
    exec_country = st.text_input("Country", placeholder="e.g. United States")
    exec_photo = st.file_uploader("Executive photo", type=["jpg", "jpeg", "png"])

st.subheader("2. LinkedIn activity data")
uploaded_excel = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

rows_per_slide = st.number_input(
    "Rows per slide (leave as-is unless you changed the template's table)",
    min_value=1, max_value=20, value=3, step=1,
)

if uploaded_excel is not None:
    if st.button("Generate PowerPoint"):
        with st.spinner("Building your presentation..."):
            try:
                photo_bytes = exec_photo.read() if exec_photo is not None else None
                output = generate_pptx(
                    TEMPLATE_PATH,
                    uploaded_excel,
                    rows_per_slide=rows_per_slide,
                    linkedin_id=linkedin_id or None,
                    exec_name=exec_name or None,
                    exec_country=exec_country or None,
                    picture_bytes=photo_bytes,
                )
                st.success("Done! Your presentation is ready.")
                fname = f"{(exec_name or 'LinkedIn_Analysis').replace(' ', '_')}_Output.pptx"
                st.download_button(
                    label="⬇️ Download PPTX",
                    data=output,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            except Exception as e:
                st.error(f"Something went wrong: {e}")
