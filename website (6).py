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
    required_columns = ["Email id", "Personal Status", "Academic Status", "Upload Status", "Payment Status"]
    missing_prev = [col for col in required_columns if col not in prev_df.columns]
    missing_curr = [col for col in required_columns if col not in curr_df.columns]
    
    if missing_prev or missing_curr:
        missing_msg = ""
        if missing_prev:
            missing_msg += f"⚠️ Missing columns in previous sheet: {missing_prev}"
        if missing_curr:
            missing_msg += f"⚠️ Missing columns in current sheet: {missing_curr}"
        st.error(missing_msg)
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

        # Define stage order for consistent display
        stage_order = ["Registered", "Personal", "Academic", "Upload", "Payment"]
        
        # Merge the sheets on "Email id" to determine stage transitions for common users
        common_users = pd.merge(
            prev_df[["Email id", "Stage"]],
            curr_df[["Email id", "Stage"]],
            on="Email id", 
            how="inner",
            suffixes=("_prev", "_curr")
        )
        
        # Create transition data with specific format requested
        transitions = []
        
        # For Registered in previous sheet (4 possible transitions)
        for dest in ["Personal", "Academic", "Upload", "Payment"]:
            count = len(common_users[(common_users["Stage_prev"] == "Registered") & 
                                     (common_users["Stage_curr"] == dest)])
            transitions.append({
                "Previous Stage": "Registered",
                "Current Stage": dest,
                "Count": count
            })
            
        # For Personal in previous sheet (3 possible transitions)
        for dest in ["Academic", "Upload", "Payment"]:
            count = len(common_users[(common_users["Stage_prev"] == "Personal") & 
                                     (common_users["Stage_curr"] == dest)])
            transitions.append({
                "Previous Stage": "Personal",
                "Current Stage": dest,
                "Count": count
            })
            
        # For Academic in previous sheet (2 possible transitions)
        for dest in ["Upload", "Payment"]:
            count = len(common_users[(common_users["Stage_prev"] == "Academic") & 
                                     (common_users["Stage_curr"] == dest)])
            transitions.append({
                "Previous Stage": "Academic",
                "Current Stage": dest,
                "Count": count
            })
            
        # For Upload in previous sheet (1 possible transition)
        count = len(common_users[(common_users["Stage_prev"] == "Upload") & 
                                 (common_users["Stage_curr"] == "Payment")])
        transitions.append({
            "Previous Stage": "Upload",
            "Current Stage": "Payment",
            "Count": count
        })
        
        # Also add users who stayed in the same stage
        for stage in stage_order:
            count = len(common_users[(common_users["Stage_prev"] == stage) & 
                                     (common_users["Stage_curr"] == stage)])
            transitions.append({
                "Previous Stage": stage,
                "Current Stage": stage,
                "Count": count
            })
        
        # Convert to DataFrame
        transition_df = pd.DataFrame(transitions)
        
        # Display the detailed transitions
        st.subheader("🔁 Stage Transitions")
        st.dataframe(transition_df, use_container_width=True)
        
        # Create pivot table for easier visualization
        pivot_df = pd.pivot_table(
            transition_df,
            values="Count",
            index=["Previous Stage"],
            columns=["Current Stage"],
            fill_value=0
        )
        
        # Ensure proper ordering of rows and columns
        pivot_df = pivot_df.reindex(stage_order, axis=0)
        pivot_cols = [col for col in stage_order if col in pivot_df.columns]
        pivot_df = pivot_df[pivot_cols]
        
        st.subheader("📊 Stage Transition Matrix")
        st.dataframe(pivot_df, use_container_width=True)

        # Provide CSV download buttons
        trans_csv = transition_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Stage Transitions (CSV)", 
            trans_csv, 
            file_name="stage_transitions.csv"
        )
        
        matrix_csv = pivot_df.reset_index().to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Stage Matrix (CSV)", 
            matrix_csv, 
            file_name="stage_matrix.csv"
        )
else:
    st.info("👈 Upload both sheets to begin analysis.")
