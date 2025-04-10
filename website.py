import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stage Shift Dashboard", layout="wide")
st.title("📊 Stage Shift Comparison Dashboard")

st.markdown("Upload the **previous** and **current** Excel sheets to analyze new registrations and status transitions.")

# Upload previous and current sheets
prev_file = st.file_uploader("Upload Previous Sheet (e.g. April 1)", type=["xlsx"], key="prev")
curr_file = st.file_uploader("Upload Current Sheet (e.g. April 2)", type=["xlsx"], key="curr")

# Safe reader for Excel files using pandas and Streamlit uploader
def safe_read_excel(uploaded_file):
    try:
        uploaded_file.seek(0)  # Reset pointer before reading
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to load Excel file: {e}")
        st.stop()

# Determine the stage from status columns
def determine_stage(row):
    if row.get("Payment_Status") == "Completed":
        return "Payment"
    elif row.get("Upload_Status") == "Completed":
        return "Upload"
    elif row.get("Academic_Status") == "Completed":
        return "Academic"
    elif row.get("Personal_Status") == "Completed":
        return "Personal"
    else:
        return "Registered"

# Main logic
if prev_file and curr_file:
    prev_df = safe_read_excel(prev_file)
    curr_df = safe_read_excel(curr_file)

    if "Email_ID" not in prev_df.columns or "Email_ID" not in curr_df.columns:
        st.error("⚠️ 'Email_ID' column must be present in both sheets.")
    else:
        # Add Stage column
        prev_df["Stage"] = prev_df.apply(determine_stage, axis=1)
        curr_df["Stage"] = curr_df.apply(determine_stage, axis=1)

        # Count total and new users
        total_prev = prev_df["Email_ID"].nunique()
        new_users_df = curr_df[~curr_df["Email_ID"].isin(prev_df["Email_ID"])]
        total_new = new_users_df["Email_ID"].nunique()

        # Summary Table
        summary_df = pd.DataFrame({
            "Metric": ["Total in Previous Sheet", "New Users in Current Sheet"],
            "Count": [total_prev, total_new]
        })

        st.subheader("📌 Registration Summary")
        st.dataframe(summary_df, use_container_width=True)

        # Stage shift analysis
        merged = pd.merge(prev_df[["Email_ID", "Stage"]],
                          curr_df[["Email_ID", "Stage"]],
                          on="Email_ID", suffixes=("_prev", "_curr"))

        stage_shift_df = merged[merged["Stage_prev"] != merged["Stage_curr"]]
        shift_counts = stage_shift_df.groupby(["Stage_prev", "Stage_curr"]).size().reset_index(name="Count")

        st.subheader("🔁 Stage Transitions")
        st.dataframe(shift_counts, use_container_width=True)

        # CSV Download
        csv = shift_counts.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Stage Shift Report (CSV)", csv, file_name="stage_shift_report.csv")
else:
    st.info("👈 Upload both sheets to begin analysis.")
