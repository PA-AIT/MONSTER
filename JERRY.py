import streamlit as st
import imaplib
import email
import re
import pandas as pd

# Streamlit app title
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user and password
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")

# Create input field for the email address to search for
search_email = st.text_input("Enter the email address to search for")

# Function to extract information from plain text content
def extract_info_from_plain_text(text_content):
    info = {
        "Name": None,
        "Email": None,
        "Phone": None,
        "Page URL": None,
        "Page Name": None
    }

    # Define patterns for different labels
    label_patterns = {
        "Name": re.compile(r'Name\s*:\s*(.*)', re.IGNORECASE),
        "Email": re.compile(r'Email\s*:\s*(.*)', re.IGNORECASE),
        "Phone": re.compile(r'Phone\s*:\s*(.*)', re.IGNORECASE),
        "Page URL": re.compile(r'Page URL\s*:\s*(.*)', re.IGNORECASE),
        "Page Name": re.compile(r'Page Name\s*:\s*(.*)', re.IGNORECASE),
    }

    for label, pattern in label_patterns.items():
        match = pattern.search(text_content)
        if match:
            info[label] = match.group(1).strip()

    return info

if st.button("Fetch and Generate Excel"):
    try:
        if not user or not password:
            st.warning("Please provide both your email address and password.")
        else:
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
            value = search_email  # Use the user-inputted email address to search
            _, data = my_mail.search(None, key, value)

            mail_id_list = data[0].split()

            info_list = []

            # Iterate through messages and extract information from plain text content
            for num in mail_id_list:
                typ, data = my_mail.fetch(num, '(RFC822)')
                msg = email.message_from_bytes(data[0][1])

                # Check for plain text content
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == 'text/plain':
                            text_content = part.get_payload(decode=True).decode('utf-8')
                            info = extract_info_from_plain_text(text_content)

                            # Extract and add the received date
                            date = msg["Date"]
                            info["Received Date"] = date

                            info_list.append(info)

            # Create a DataFrame from the info_list
            df = pd.DataFrame(info_list)

            # Generate the Excel file
            st.write("Data extracted from emails:")
            st.write(df)

            if st.button("Download Excel File"):
                excel_file = df.to_excel('EXPO_leads.xlsx', index=False, engine='openpyxl')
                if excel_file:
                    with open('EXPO_leads.xlsx', 'rb') as file:
                        st.download_button(
                            label="Click to download Excel file",
                            data=file,
                            key='download-excel'
                        )

            st.success("Excel file has been generated and is ready for download.")

    except Exception as e:
        st.error(f"Error: {e}")
