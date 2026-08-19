import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="TGNA PDF Extractor", layout="wide")

st.title("TGNA PDF Extractor")

uploaded_files = st.file_uploader(
    "Upload TGNA Approval PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if st.button("Generate Excel"):

    if not uploaded_files:
        st.warning("Please upload PDF files.")
        st.stop()

    records = []

    for uploaded_file in uploaded_files:

        try:

            text = ""

            with pdfplumber.open(uploaded_file) as pdf:

                for page in pdf.pages:

                    page_text = page.extract_text(
                        x_tolerance=2,
                        y_tolerance=2
                    )

                    if page_text:
                        text += page_text + "\n"

            text = text.replace("\t", " ")

            record = {
                "PDF Name": uploaded_file.name,
                "Nodal RLDC": "",
                "Application Number": "",
                "Acceptance Date": "",
                "Name Of Applicant": "",
                "Injection Entity/State/Region": "",
                "Drawee Entity/State/Region": "",
                "Green Tag": "",
                "Is T-GNA-RE application": "",
                "From": "",
                "To": "",
                "Total MWh(To be scheduled)": "",
                "Transmission Charges": "",
                "Operating Charges": ""
            }

            # Nodal RLDC
            m = re.search(
                r'Nodal\s*RLDC:\s*([A-Z]+)',
                text,
                re.IGNORECASE
            )

            if m:
                record["Nodal RLDC"] = m.group(1)

            # Application Number
            m = re.search(
                r'Application\s*Number:\s*(\d+)',
                text,
                re.IGNORECASE
            )

            if m:
                record["Application Number"] = m.group(1)

            # Acceptance Date
            m = re.search(
                r'Acceptance\s*Date:\s*(\d{2}/\d{2}/\d{4})',
                text,
                re.IGNORECASE
            )

            if m:
                record["Acceptance Date"] = m.group(1)

            # Acceptance Number
            m = re.search(
                r'Acceptance\s*No:\s*(\d{2}/\d{2}/\d{4})',
                text,
                re.IGNORECASE
            )

            if m:
                record["Acceptance No"] = m.group(1)

            # Name Of Applicant
            m = re.search(
                r'Name\s*Of\s*Applicant:\s*(.*?)\s*2\.',
                text,
                re.IGNORECASE | re.DOTALL
            )

            if m:
                record["Name Of Applicant"] = re.sub(
                    r"\s+",
                    " ",
                    m.group(1)
                ).strip()

            # Injection Entity
            inj_match = re.search(
                r'Injection\s*Entity/State/Region:\s*(.*?/(?:WR|NR|SR|ER|NER))',
                text,
                re.IGNORECASE | re.DOTALL
            )

            if inj_match:

                injection = re.sub(
                    r"\s+",
                    " ",
                    inj_match.group(1)
                ).strip()

                record["Injection Entity/State/Region"] = injection

                remaining = text[inj_match.end():]

                dr_match = re.search(
                    r'(.*?/(?:WR|NR|SR|ER|NER))',
                    remaining,
                    re.IGNORECASE | re.DOTALL
                )

                if dr_match:

                    drawee = dr_match.group(1)

                    drawee = re.sub(
                        r'^\s*3\.\s*Drawee\s*Entity/State/Region:\s*',
                        '',
                        drawee,
                        flags=re.IGNORECASE
                    )

                    drawee = re.sub(
                        r'\s+',
                        ' ',
                        drawee
                    ).strip()

                    record["Drawee Entity/State/Region"] = drawee

            # Green Tag
            m = re.search(
                r'Green\s*Tag:\s*(Yes|No)',
                text,
                re.IGNORECASE
            )

            if m:
                record["Green Tag"] = m.group(1)

            # TGNA RE
            m = re.search(
                r'Is\s*T-GNA-RE\s*application:\s*(Yes|No)',
                text,
                re.IGNORECASE
            )

            if m:
                record["Is T-GNA-RE application"] = m.group(1)

            # From and To Date
            m = re.search(
                r'Open\s*Access\s*Scheduling\s*Accepted:.*?'
                r'(\d{2}/\d{2}/\d{4})\s+'
                r'(\d{2}/\d{2}/\d{4})',
                text,
                re.IGNORECASE | re.DOTALL
            )

            if m:
                record["From"] = m.group(1)
                record["To"] = m.group(2)

            # Total MWh(To be scheduled)
            m = re.search(
                r'Total\s*MWh\s*\(To\s*be\s*scheduled\)\s*([0-9,]+)',
                text,
                re.IGNORECASE
            )

            if m:
                record["Total MWh(To be scheduled)"] = m.group(1)

            # Transmission Charges
            m = re.search(
                r'Total\s*Of\s*\(?1\)?\s*[₹Rs.\s]*([0-9,]+)',
                text,
                re.IGNORECASE
            )

            if m:
                record["Transmission Charges"] = m.group(1)

            # Non-Refundable Application Fee
            m = re.search(
                r'Non-Refundable\s*Application\s*Fee.*?₹\s*([0-9,]+)',
                text,
                re.IGNORECASE | re.DOTALL
            )

            if m:
                record["Operating Charges"] = m.group(1)

            records.append(record)

        except Exception as e:

            st.error(
                f"Error processing {uploaded_file.name}: {str(e)}"
            )

    df = pd.DataFrame(records)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df.to_excel(
            writer,
            sheet_name="TGNA Data",
            index=False
        )

        ws = writer.sheets["TGNA Data"]

        for column in ws.columns:

            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:

                try:
                    if cell.value:
                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )
                except:
                    pass

            ws.column_dimensions[column_letter].width = min(
                max_length + 5,
                60
            )

    output.seek(0)

    st.success("Excel Generated Successfully")

    st.dataframe(df)

    st.download_button(
        label="Download TGNA_Consolidated.xlsx",
        data=output,
        file_name="TGNA_Consolidated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
