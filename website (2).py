import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stage Shift Dashboard", layout="wide")
st.title("📊 Stage Shift Comparison Dashboard")

st.markdown("Upload the **previous** and **current** Excel sheets to analyze new registrations and status transitions.")

# Upload previous and current sheets
prev_file = st.file_uploader("Upload Previous Sheet (e.g. April 1)", type=["xlsx"], key="prev")
curr_file = st.file_uploader("Upload Current Sheet (e.g. April 2)", type=["xlsx"], key="curr")

# Safe reader: reset the file pointer and read using pandas.read_excel
def safe_read_excel(uploaded_file):
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to load Excel file: {e}")
        st.stop()

# Determine the user's stage based on status columns (note the spaces in column names)
def determine_stage(row):
    if row.get("Payment Status") == "Completed":
        return "Payment"
    elif row.get("Upload Status") == "Completed":
        return "Upload"
    elif row.get("Academic Status") == "Completed":
        return "Academic"
    elif row.get("Personal Status") == "Completed":
        return "Personal"
    else:
        return "Registered"

if prev_file and curr_file:
    prev_df = safe_read_excel(prev_file)
    curr_df = safe_read_excel(curr_file)

    # Check that the sheets have the expected "Email" column
    if "Email" not in prev_df.columns or "Email" not in curr_df.columns:
        st.error("⚠️ 'Email' column must be present in both sheets.")
    else:
        # Determine stage for each row
        prev_df["Stage"] = prev_df.apply(determine_stage, axis=1)
        curr_df["Stage"] = curr_df.apply(determine_stage, axis=1)

        # Total registered users in the previous sheet (based on "Email")
        total_prev = prev_df["Email"].nunique()
        # New users in the current sheet (not present in previous)
        new_users_df = curr_df[~curr_df["Email"].isin(prev_df["Email"])]
        total_new = new_users_df["Email"].nunique()

        # Create a summary table with registration data
        summary_df = pd.DataFrame({
            "Metric": ["Total in Previous Sheet", "New Users in Current Sheet"],
            "Count": [total_prev, total_new]
        })

        st.subheader("📌 Registration Summary")
        st.dataframe(summary_df, use_container_width=True)

        # Merge the sheets on "Email" to compare stage changes
        merged = pd.merge(prev_df[["Email", "Stage"]],
                          curr_df[["Email", "Stage"]],
                          on="Email", suffixes=("_prev", "_curr"))

        # Identify users whose stage has changed
        stage_shift_df = merged[merged["Stage_prev"] != merged["Stage_curr"]]
        shift_counts = stage_shift_df.groupby(["Stage_prev", "Stage_curr"]).size().reset_index(name="Count")

        st.subheader("🔁 Stage Transitions")
        st.dataframe(shift_counts, use_container_width=True)

        # Provide a download button for the stage shift report as CSV
        csv = shift_counts.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Stage Shift Report (CSV)", csv, file_name="stage_shift_report.csv")
else:
    st.info("👈 Upload both sheets to begin analysis.")
