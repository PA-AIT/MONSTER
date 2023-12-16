import streamlit as st
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd
from io import BytesIO

# Streamlit app title
st.header("Developed by MKSSS-AIT")
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user, password, and mail address
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")
mail_address = st.text_input("Enter the mail address from which you want to extract information")

# Function to extract information from HTML content
def extract_info_from_html(html_content):
    # ... (unchanged)

if st.button("Fetch and Generate Excel"):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using user and password
        my_mail.login(user, password)

        # Select the Inbox to fetch messages
        my_mail.select('inbox')

        # Define the key and value for email search
        key = 'FROM'
        value = mail_address  # Use the specified mail address
        _, data = my_mail.search(None, key, value)

        mail_id_list = data[0].split()

        info_list = []

        # Iterate through messages and extract information from HTML content
        for num in mail_id_list:
            # ... (unchanged)

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Display the data in the Streamlit app
        st.write("Data extracted from emails:")
        st.write(df)

        # Download the DataFrame as an Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        output.seek(0)
        st.write("Downloading Excel file...")
        st.download_button(
            label="Download Excel File",
            data=output,
            key="download_excel",
            on_click=None,
            file_name="EXPO_leads.xlsx"  # Specify the file name
        )

    except Exception as e:
        st.error(f"Error: {e}")
