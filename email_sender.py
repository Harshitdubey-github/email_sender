import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Email sending function
def send_email(subject, body, to_email, from_email, password, bcc_email=None):
    """Send an email."""
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    if bcc_email:
        msg['Bcc'] = bcc_email

    # Attach the body
    msg.attach(MIMEText(body, 'plain'))

    try:
        smtp_server = "send.smtp.com"
        smtp_port = 2525
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)
        server.quit()

        logger.info(f"Email sent successfully to {to_email}!")
        return True
    except smtplib.SMTPException as e:
        logger.error(f"Failed to send email to {to_email}: {e}")
        return False

# Streamlit app
def main():
    st.title("ScribeEMR Bulk Email Sender")

    # Email account credentials
    st.header("Email Account Credentials")
    from_email = st.text_input("Your Email Address")
    password = st.text_input("Your Email Password", type="password")

    # BCC Email (optional)
    bcc_email = st.text_input("BCC Email Address (Optional)")

    # Upload Excel file
    st.header("Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = {'first name', 'last name', 'email id'}
            if not required_columns.issubset(df.columns.str.lower()):
                st.error(f"Excel file must contain the following columns: {required_columns}")
                return
            else:
                st.success("Excel file uploaded successfully!")
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return

    # Email content
    st.header("Email Content")
    subject = st.text_input("Email Subject")
    body = st.text_area("Email Body (will start with 'Dear {first name}')")

    # Send emails
    if st.button("Send Emails"):
        if not uploaded_file:
            st.error("Please upload an Excel file.")
        elif not from_email or not password:
            st.error("Please provide your email address and password.")
        elif not subject or not body:
            st.error("Please provide both a subject and a body for the email.")
        else:
            df = pd.read_excel(uploaded_file)
            # Normalize column names to lower case
            df.columns = df.columns.str.lower()
            sent_count = 0
            for index, row in df.iterrows():
                first_name = row['first name']
                to_email = row['email id']
                personalized_body = f"Dear {first_name},\n\n{body}"
                success = send_email(
                    subject,
                    personalized_body,
                    to_email,
                    from_email,
                    password,
                    bcc_email=bcc_email if bcc_email else None
                )
                if success:
                    sent_count += 1
            st.success(f"Emails sent successfully to {sent_count} recipients!")

if __name__ == "__main__":
    main()
