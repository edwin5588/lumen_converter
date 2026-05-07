import tempfile
import os
import streamlit as st
import conversion_script

st.title("Lumen Converter")
st.write("Upload plate reader `.xls` or `.xlsx` files to convert them into formatted plates.")

uploaded_files = st.file_uploader("Choose files", type=["xls", "xlsx"], accept_multiple_files=True)

for uploaded in uploaded_files:
    suffix = os.path.splitext(uploaded.name)[1]
    out_path = None
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
        tmp_in.write(uploaded.read())
        tmp_in_path = tmp_in.name

    try:
        with st.spinner(f"Converting {uploaded.name}..."):
            out_path = conversion_script.convert(tmp_in_path)

        with open(out_path, "rb") as f:
            out_bytes = f.read()

        out_name = os.path.splitext(uploaded.name)[0] + "_CONVERTED_plates.xlsx"
        st.success(f"{uploaded.name} converted.")
        st.download_button(
            label=f"Download {out_name}",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=uploaded.name,
        )
    except Exception as e:
        st.error(f"{uploaded.name} failed: {e}")
    finally:
        os.unlink(tmp_in_path)
        if out_path and os.path.exists(out_path):
            os.unlink(out_path)
