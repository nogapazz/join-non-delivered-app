import streamlit as st
import pandas as pd
from io import BytesIO

# Load the "All Contacts" file from Streamlit secrets storage
import base64
from io import BytesIO

@st.cache_data
def load_contacts():
    encoded_data = st.secrets["all_contacts"]["content"]  # Retrieve Base64 data
    decoded_bytes = base64.b64decode(encoded_data)  # Decode it
    return pd.read_excel(BytesIO(decoded_bytes))  # Load it as an Excel file

# App Title
st.title("Join Non-Delivered Emails with Account & CS Owner Details")

# File upload for Non-Deliverable List
st.subheader("Upload the Non-Deliverable List (Excel File)")
non_deliverable_file = st.file_uploader("Upload your file", type=["xlsx"], key="file1")

if non_deliverable_file:
    # Load the uploaded file
    non_deliverable_df = pd.read_excel(non_deliverable_file)

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
    st.dataframe(joined_df)

    # Convert to Excel for download
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        joined_df.to_excel(writer, index=False)
    buffer.seek(0)

    # Provide download button
    st.subheader("Download the Joined File")
    st.download_button(
        label="Download Joined File",
        data=buffer,
        file_name="joined_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


