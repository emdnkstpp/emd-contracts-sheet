import streamlit as st
import pandas as pd
import datetime
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="Cloud Excel Dashboard", layout="wide", page_icon="☁️")

st.title("☁️ Contract Expiry Web Dashboard")
st.markdown("This dashboard reads your predetermined `contracts.xlsx` file in real-time. Use the controls below to safely dispatch batched expiry emails via your Server Email Address.")

# Sidebar Settings
with st.sidebar:
    st.header("⚙️ Mailer Settings")
    days_notice = st.number_input("Days Notice Warning Threshold:", min_value=0, value=180)
    st.markdown("---")
    force_send = st.checkbox("Force Send All Due (Ignore 7-Day Rule)", value=False, help="Check this to override the 7-day minimum interval between emails.")
    st.markdown("---")
    
    st.markdown("**SMTP Server Email Credentials**")
    try:
        secret_email = st.secrets["EMAIL_USERNAME"] if "EMAIL_USERNAME" in st.secrets else "emdnkstppcontracts@outlook.com"
        secret_pass = st.secrets["EMAIL_PASSWORD"] if "EMAIL_PASSWORD" in st.secrets else ""
        secret_server = st.secrets["SMTP_SERVER"] if "SMTP_SERVER" in st.secrets else "smtp.office365.com"
        secret_port = st.secrets["SMTP_PORT"] if "SMTP_PORT" in st.secrets else 587
    except Exception:
        secret_email = "emdnkstppcontracts@outlook.com"
        secret_pass = ""
        secret_server = "smtp.office365.com"
        secret_port = 587
        
    smtp_server = st.text_input("SMTP Server", value=secret_server)
    smtp_port = st.number_input("SMTP Port", value=secret_port)
    sender_email = st.text_input("Sender Email Username", value=secret_email)
    sender_password = st.text_input("App Password", type="password", value=secret_pass, help="Never hardcode your password in GitHub! Use Streamlit Secrets.")
    
    st.info("💡 **Security Note**: When deploying to **Streamlit Community Cloud via GitHub**, do not type your password into your code! Save it securely under your App Settings -> Secrets.")

# File Loading Logic
default_file = "contracts.xlsx"
df = None
file_loaded = False

if os.path.exists(default_file):
    try:
        df = pd.read_excel(default_file)
        df.columns = df.columns.astype(str).str.strip()
        file_loaded = True
        st.success(f"Successfully auto-loaded `{default_file}` from the server.")
    except Exception as e:
        st.error(f"Failed to read `{default_file}`: {e}")
else:
    st.warning(f"Could not find `{default_file}` in the current directory. Please make sure it's uploaded alongside this script.")

if file_loaded and df is not None:
    # Handle Expiry Date column dynamically
    expiry_col = None
    for col in df.columns:
        if col.lower() in ['expiry date', 'expiry_date', 'expiry']:
            expiry_col = col
            break
            
    if expiry_col:
        today = datetime.datetime.now()
        temp_dates = pd.to_datetime(df[expiry_col], dayfirst=True, errors='coerce')
        df.insert(len(df.columns), 'Days to Expiry (Calculated)', temp_dates.apply(lambda x: None if pd.isna(x) else (x - today).days))
        
    # Convert everything to string for safe rendering in the browser
    st.dataframe(df.astype(str), use_container_width=True, height=500)
    
    st.markdown("### 📧 Dispatch Server Emails")
    
    if st.button("Attempt to Send Batched Emails", type="primary"):
        required = ["Contract Name", "Expiry Date", "Email Address", "Person Name"]
        missing = [col for col in required if col not in df.columns]
        
        if missing:
            st.error(f"Cannot reliably send emails. Your data is missing the following required column names: **{', '.join(missing)}**")
        elif not sender_email or not sender_password:
            st.error("Please enter your **Sender Email** and **App Password** in the sidebar to securely connect to the SMTP Server.")
        else:
            with st.spinner('Connecting to SMTP Server and processing batched emails...'):
                sent_count = 0
                today = datetime.datetime.now()
                
                try:
                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()
                    server.login(sender_email, sender_password)
                    
                    emails_to_send = {}
                    
                    for index, row in df.iterrows():
                        if pd.isna(row.get('Email Address')) or pd.isna(row.get('Expiry Date')):
                            continue
                            
                        try:
                            exp_date = pd.to_datetime(row['Expiry Date'], errors='coerce')
                        except Exception:
                            continue
                            
                        if pd.isna(exp_date):
                            continue
                            
                        days_left = (exp_date - today).days
                        
                        if 0 <= days_left <= days_notice:
                            last_emailed = row.get('Last Emailed Date', pd.NaT)
                            needs_email = False
                            
                            if force_send:
                                needs_email = True
                            elif pd.isna(last_emailed):
                                needs_email = True
                            else:
                                try:
                                    if (today - pd.to_datetime(last_emailed)).days >= 7:
                                        needs_email = True
                                except Exception:
                                    needs_email = True
                                    
                            if needs_email:
                                email = str(row['Email Address']).strip()
                                if email not in emails_to_send:
                                    emails_to_send[email] = {
                                        'person_name': str(row.get('Person Name', 'Colleague')),
                                        'contracts': [],
                                        'indices': []
                                    }
                                emails_to_send[email]['contracts'].append({
                                    'name': str(row.get('Contract Name', 'Unknown')),
                                    'expiry': exp_date,
                                    'days_left': days_left
                                })
                                emails_to_send[email]['indices'].append(index)
                                
                    for email, data in emails_to_send.items():
                        try:
                            msg = MIMEMultipart()
                            msg['From'] = sender_email
                            msg['To'] = email
                            msg['Subject'] = f"Reminder: {len(data['contracts'])} Contract(s) Expiring Soon"
                            
                            rows_html = ""
                            for c in data['contracts']:
                                rows_html += f"<tr><td style='border: 1px solid #ddd; padding: 8px;'>{c['name']}</td>"
                                rows_html += f"<td style='border: 1px solid #ddd; padding: 8px; color: #d35400;'>{c['expiry'].strftime('%Y-%m-%d')}</td>"
                                rows_html += f"<td style='border: 1px solid #ddd; padding: 8px;'>{c['days_left']}</td></tr>"
                                
                            table_html = f"<table style='border-collapse: collapse; width: 100%;'><tr><th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f4f6f9;'>Contract Name</th><th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f4f6f9;'>Expiry Date</th><th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f4f6f9;'>Days Left</th></tr>{rows_html}</table>"
                            
                            html_body = f'''
                            <html>
                            <body style="font-family: Arial, sans-serif;">
                                <p>Dear {data['person_name']},</p>
                                <p>This is an automated dashboard reminder that the following contract(s) assigned to you are set to expire soon:</p>
                                {table_html}
                                <br>
                                <p>Please take the necessary actions to renew or close these contracts as early as possible.</p>
                                <br>
                                <p>Best regards,<br>Contract Management Web Dashboard</p>
                            </body>
                            </html>
                            '''
                            
                            msg.attach(MIMEText(html_body, 'html'))
                            
                            server.send_message(msg)
                            
                            # Safely attempt to update df for saving
                            if 'Last Emailed Date' not in df.columns:
                                df['Last Emailed Date'] = pd.NaT
                            for idx in data['indices']:
                                df.at[idx, 'Last Emailed Date'] = today
                                
                            sent_count += 1
                            
                        except Exception as e:
                            st.error(f"Failed to send email to {email}: {e}")
                            
                    server.quit()
                    
                    if sent_count > 0:
                        st.balloons()
                        st.success(f"**Process complete!** Successfully sent **{sent_count}** batched expiry reminder emails using {sender_email}.")
                        
                        # Save the changes locally if possible
                        try:
                            df.to_excel(default_file, index=False)
                            st.info("Successfully saved 'Last Emailed Date' updates back to the Excel file.")
                        except Exception as e:
                            st.error(f"Could not save the Excel file to update 'Last Emailed Date' tracking: {e}")
                    else:
                        st.warning("No due emails were sent. Either all expiring contracts have been emailed within the last 7 days, or there are no contracts matching the criteria.")
                        
                except Exception as e:
                    st.error(f"Failed to connect to SMTP Server at {smtp_server}:{smtp_port}. Check your email/password and server settings. Error: {e}")

if __name__ == "__main__":
    import sys
    from streamlit.web import cli as stcli
    
    # This automatically boots up the Streamlit server if you click "Run" in VS Code!
    sys.argv = ["streamlit", "run", sys.argv[0]]
    sys.exit(stcli.main())

