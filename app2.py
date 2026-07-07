import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="Question Answer Parser",
    page_icon="📄",
    layout="wide"
)

st.title("📄 แยกข้อมูลจาก question_answer")

uploaded_file = st.file_uploader(
    "อัปโหลดไฟล์ Excel",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    # อ่านไฟล์
    df = pd.read_excel(uploaded_file)

    st.success(f"อ่านข้อมูลสำเร็จ ({len(df)} รายการ)")

    st.subheader("ข้อมูลต้นฉบับ")
    st.dataframe(df, use_container_width=True)

    # ตรวจสอบคอลัมน์
    if "question_answer" not in df.columns:
        st.error("ไม่พบคอลัมน์ question_answer")
        st.stop()

    # ฟังก์ชันแยกคำถาม-คำตอบ
    def parse_question_answer(text):
        result = {}

        if pd.isna(text):
            return result

        pattern = r"Q\d+\s*(.*?)\nA\d+\s*(.*?)(?=\nQ\d+|\Z)"
        matches = re.findall(pattern, str(text), flags=re.S)

        for question, answer in matches:
            result[question.strip()] = answer.strip()

        return result

    # แยกข้อมูล
    parsed = df["question_answer"].apply(parse_question_answer)
    parsed_df = pd.DataFrame(parsed.tolist())

    # รวมข้อมูล
    result = pd.concat(
        [df.drop(columns=["question_answer"]), parsed_df],
        axis=1
    )

    st.subheader("ผลลัพธ์")
    st.dataframe(result, use_container_width=True)

    # Export Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Result")

    output.seek(0)

    st.download_button(
        label="📥 ดาวน์โหลด Excel",
        data=output,
        file_name="result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )