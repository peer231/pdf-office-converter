import streamlit as st

from converter import (
    extract_text_from_pdf,
    create_word,
    create_csv
)


st.set_page_config(
    page_title="Annex-C PDF Converter",
    layout="centered"
)


st.title("Annex-C PDF to Editable Word Converter")

st.write(
    "Upload the Annex-C PDF and convert it into editable Word and CSV format."
)


uploaded_file = st.file_uploader(
    "Choose PDF file",
    type=["pdf"]
)


if uploaded_file is not None:

    st.success(
        f"Uploaded: {uploaded_file.name}"
    )

    pdf_bytes = uploaded_file.getvalue()

    if st.button(
        "Convert PDF",
        type="primary"
    ):

        try:

            with st.spinner(
                "Converting PDF... Please wait."
            ):

                pages = extract_text_from_pdf(
                    pdf_bytes
                )

                word_file = create_word(
                    pdf_bytes
                )

                csv_file = create_csv(
                    pages
                )

            st.success(
                "Conversion completed successfully!"
            )

            original_name = uploaded_file.name.rsplit(".", 1)[0]

            st.download_button(
                label="Download Editable Word",
                data=word_file,
                file_name=f"{original_name}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            )

            st.download_button(
                label="Download CSV",
                data=csv_file,
                file_name=f"{original_name}.csv",
                mime="text/csv"
            )

            st.subheader(
                "Extracted Text Preview"
            )

            for page in pages:

                with st.expander(
                    f"Page {page['page']}"
                ):

                    st.text(
                        page["text"]
                    )

        except Exception as error:

            st.error(
                f"Conversion failed: {error}"
            )