import pandas as pd
import streamlit as st

# Monkey-patch openpyxl's Fill class to ignore fill errors
try:
    from openpyxl.styles.fills import Fill
    orig_init = Fill.__init__

    def safe_init(self, *args, **kwargs):
        try:
            orig_init(self, *args, **kwargs)
        except Exception:
            pass

    Fill.__init__ = safe_init
except Exception as e:
    st.error(f"Monkey-patch failed: {e}")
    st.stop()

st.set_page_config(page_title="Stage Shift Dashboard", layout="wide")
st.title("📊 Stage Shift Comparison Dashboard")
st.markdown("Upload the **previous** and **current** Excel sheets to analyze new registrations and status transitions.")

# Upload previous and current sheets
prev_file = st.file_uploader("Upload Previous Sheet (e.g. April 1)", type=["xlsx"], key="prev")
curr_file = st.file_uploader("Upload Current Sheet (e.g. April 2)", type=["xlsx"], key="curr")

# Safe reader for Excel files using pandas.read_excel; ensure file pointer is reset
def safe_read_excel(uploaded_file):
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to load Excel file: {e}")
        st.stop()

# Determine the stage from status columns
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

# Main logic
if prev_file and curr_file:
    prev_df = safe_read_excel(prev_file)
    curr_df = safe_read_excel(curr_file)

    # Check that the sheets have the correct column name
    if "Email id" not in prev_df.columns or "Email id" not in curr_df.columns:
        st.error("⚠️ 'Email id' column must be present in both sheets.")
    else:
        # Add Stage column based on the status columns
        prev_df["Stage"] = prev_df.apply(determine_stage, axis=1)
        curr_df["Stage"] = curr_df.apply(determine_stage, axis=1)

        # Count total users in previous sheet and new users in current sheet
        total_prev = prev_df["Email id"].nunique()
        new_users_df = curr_df[~curr_df["Email id"].isin(prev_df["Email id"])]
        total_new = new_users_df["Email id"].nunique()

        # Build registration summary table
        summary_df = pd.DataFrame({
            "Metric": ["Total in Previous Sheet", "New Users in Current Sheet"],
            "Count": [total_prev, total_new]
        })

        st.subheader("📌 Registration Summary")
        st.dataframe(summary_df, use_container_width=True)

        # Merge the sheets on "Email id" to determine stage transitions
        merged = pd.merge(prev_df[["Email id", "Stage"]],
                          curr_df[["Email id", "Stage"]],
                          on="Email id", suffixes=("_prev", "_curr"))

        # Filter to users whose stage has changed
        stage_shift_df = merged[merged["Stage_prev"] != merged["Stage_curr"]]
        shift_counts = stage_shift_df.groupby(["Stage_prev", "Stage_curr"]).size().reset_index(name="Count")

        st.subheader("🔁 Stage Transitions")
        st.dataframe(shift_counts, use_container_width=True)

        # Provide a CSV download button for the stage shift report
        csv = shift_counts.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Stage Shift Report (CSV)", csv, file_name="stage_shift_report.csv")
else:
    st.info("👈 Upload both sheets to begin analysis.")
