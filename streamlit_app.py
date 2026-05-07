import tempfile
import os
import streamlit as st
import conversion_script

st.title("Lumen Converter")
st.write("Upload a plate reader `.xls` or `.xlsx` file to convert it into formatted plates.")

uploaded = st.file_uploader("Choose a file", type=["xls", "xlsx"])

if uploaded is not None:
    suffix = os.path.splitext(uploaded.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
        tmp_in.write(uploaded.read())
        tmp_in_path = tmp_in.name

    try:
        with st.spinner("Converting..."):
            out_path = conversion_script.convert(tmp_in_path)

        with open(out_path, "rb") as f:
            out_bytes = f.read()

        out_name = os.path.splitext(uploaded.name)[0] + "CONVERTED_plates.xlsx"
        st.success("Conversion complete.")
        st.download_button(
            label="Download converted file",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Conversion failed: {e}")
    finally:
        os.unlink(tmp_in_path)
        if "out_path" in dir() and os.path.exists(out_path):
            os.unlink(out_path)
