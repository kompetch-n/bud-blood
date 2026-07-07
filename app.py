import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pymongo import MongoClient
from datetime import datetime
from pymongo import UpdateOne

st.set_page_config(
    page_title="บริจาคโลหิต โรงพยาบาลกรุงเทพอุดร",
    page_icon="📄",
    layout="wide"
)

# =============================
# MongoDB
# =============================
MONGO_URI = st.secrets["MONGO_URI"]  # เก็บไว้ใน secrets.toml

client = MongoClient(MONGO_URI)

db = client["hospital_forms"]          # Database
collection = db["submissions"]         # Collection

# ===========================================
# แสดงข้อมูลจาก MongoDB
# ===========================================

st.header("📋 บริจาคโลหิต โรงพยาบาลกรุงเทพอุดร")

# ดึงข้อมูล
docs = list(collection.find().sort("updated_at", -1))

if docs:

    mongo_df = pd.DataFrame(docs)

    if "_id" in mongo_df.columns:
        mongo_df["_id"] = mongo_df["_id"].astype(str)

    # คอลัมน์ที่ต้องการ Export
    export_columns = [
        "submission_id",
        "submitted_at",
        "ชื่อ - นามสกุล",
        "เบอร์โทรศัพท์",
        "แผนก",
        "หมายเหตุ"
    ]

    # เลือกเฉพาะคอลัมน์ที่มีอยู่จริง
    export_df = mongo_df.reindex(columns=export_columns)

    # Export Excel
    output2 = BytesIO()

    with pd.ExcelWriter(output2, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Blood Donation")

    output2.seek(0)

    # ปุ่มด้านบน
    col1, col2, col3 = st.columns([1, 2, 7])

    with col1:
        st.button("🔄 Refresh")

    with col2:
        st.download_button(
            "📥 ดาวน์โหลด Excel",
            data=output2,
            file_name="mongodb_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.write(f"จำนวนทั้งหมด {len(mongo_df):,} รายการ")

    st.dataframe(
        mongo_df,
        use_container_width=True,
        height=500
    )

else:
    st.info("ยังไม่มีข้อมูลใน MongoDB")



uploaded_file = st.file_uploader(
    "อัปโหลดไฟล์ Excel",
    type=["xlsx", "xls"]
)

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.success(f"พบข้อมูล {len(df)} รายการ")

    if "question_answer" not in df.columns:
        st.error("ไม่พบคอลัมน์ question_answer")
        st.stop()

    def parse_question_answer(text):
        result = {}

        if pd.isna(text):
            return result

        pattern = r"Q\d+\s*(.*?)\nA\d+\s*(.*?)(?=\nQ\d+|\Z)"
        matches = re.findall(pattern, str(text), flags=re.S)

        for q, a in matches:
            result[q.strip()] = a.strip()

        return result

    parsed = df["question_answer"].apply(parse_question_answer)
    parsed_df = pd.DataFrame(parsed.tolist())

    result = pd.concat(
        [df.drop(columns=["question_answer"]), parsed_df],
        axis=1
    )

    st.subheader("ข้อมูลที่แยกแล้ว")
    st.dataframe(result, use_container_width=True)

    # -------------------------
    # Export Excel
    # -------------------------
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False)

    output.seek(0)

    st.download_button(
        "📥 ดาวน์โหลด Excel",
        output,
        "result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # -------------------------
    # Save MongoDB (Upsert)
    # -------------------------
    if st.button("💾 บันทึกลง MongoDB"):

        records = result.fillna("").to_dict("records")

        operations = []

        now = datetime.utcnow()

        for doc in records:

            doc["updated_at"] = now

            operations.append(
                UpdateOne(
                    {"submission_id": doc["submission_id"]},   # เช็คข้อมูลซ้ำ
                    {
                        "$set": doc,
                        "$setOnInsert": {
                            "created_at": now
                        }
                    },
                    upsert=True
                )
            )

        if operations:

            bulk_result = collection.bulk_write(operations)

            inserted = bulk_result.upserted_count
            modified = bulk_result.modified_count

            st.success(
                f"""บันทึกเรียบร้อย

    เพิ่มใหม่ : {inserted} รายการ

    อัปเดต : {modified} รายการ"""
            )