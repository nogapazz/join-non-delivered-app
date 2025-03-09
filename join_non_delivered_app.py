import streamlit as st
import base64
from io import BytesIO
import pandas as pd

# Load contacts function (securely retrieves the secret file)
@st.cache_data
def load_contacts():
    encoded_data = st.secrets["all_contacts"]["content"]
    decoded_bytes = base64.b64decode(encoded_data)
    return pd.read_excel(BytesIO(decoded_bytes), engine='openpyxl')

# Function to load uploaded file in multiple formats
def load_uploaded_file(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    elif uploaded_file.name.endswith('.xlsx'):
        return pd.read_excel(uploaded_file, engine='openpyxl')
    else:
        st.error("Unsupported file format. Please upload a .csv or .xlsx file.")
        st.stop()

# Set background image function
def set_background(image_path):
    with open(image_path, "rb") as f:
        encoded_image = base64.b64encode(f.read()).decode()
    
    bg_css = f"""
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{encoded_image}");
        background-size: cover;
    }}
    </style>
    """
    st.markdown(bg_css, unsafe_allow_html=True)

# Call the background function
background_image_path = "assets/background.png"
set_background(background_image_path)

# App Title
st.title("Join Non-Delivered Emails with Account & CS Owner Details")

# File upload for Non-Deliverable List
st.subheader("Upload the Non-Deliverable List (Excel or CSV)")
non_deliverable_file = st.file_uploader("Upload your file", type=["csv", "xlsx"], key="file1")

if non_deliverable_file:
    # Load the uploaded file
    non_deliverable_df = load_uploaded_file(non_deliverable_file)

    # Load the "All Contacts" file from storage
    all_contacts_df = load_contacts()

    # Perform the join operation (Left Join)
    joined_df = non_deliverable_df.merge(
        all_contacts_df,
        left_on="Recipient",
        right_on="Email",
        how="left"
    )

    # Drop the Email column since we already have Recipient
    joined_df = joined_df.drop(columns=["Email"], errors="ignore")

    # Display success message and preview
    st.success("Files joined successfully!")
    st.write("Preview of the joined file:")
    st.dataframe(joined_df.head(10))  # Show only 10 rows for privacy

    # Convert joined data to CSV and XLSX
    csv_buffer = joined_df.to_csv(index=False).encode('utf-8')

    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        joined_df.to_excel(writer, index=False)
    excel_buffer.seek(0)

    # Provide download buttons
    st.subheader("Download Options")
    st.download_button(
        label="📥 Download as CSV",
        data=csv_buffer,
        file_name="joined_file.csv",
        mime="text/csv"
    )

    st.download_button(
        label="📥 Download as Excel (XLSX)",
        data=excel_buffer,
        file_name="joined_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Optional: Open in Google Sheets (via a public upload link)
    google_sheets_link = "https://docs.google.com/spreadsheets/u/0/create"
    st.markdown(f"[📄 Open in Google Sheets]( {google_sheets_link} )", unsafe_allow_html=True)
