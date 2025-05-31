import streamlit as st
import pandas as pd
import gspread
import qrcode
import io
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
import os
import json
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# ===== Gmail 인증 =====
def get_gmail_service():
    SCOPES = ['https://www.googleapis.com/auth/gmail.send']
    creds = Credentials.from_authorized_user_info(st.session_state['credentials'], SCOPES)
    service = build('gmail', 'v1', credentials=creds)
    return service

def save_credentials():
    oauth_info = st.secrets["oauth"]
    redirect_uri = st.secrets["redirect_uri"]

    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": oauth_info["client_id"],
                "project_id": oauth_info["project_id"],
                "auth_uri": oauth_info["auth_uri"],
                "token_uri": oauth_info["token_uri"],
                "auth_provider_x509_cert_url": oauth_info["auth_provider_x509_cert_url"],
                "client_secret": oauth_info["client_secret"],
                "redirect_uris": oauth_info["redirect_uris"]
            }
        },
        scopes=[
            'https://www.googleapis.com/auth/gmail.send',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ],
        redirect_uri=redirect_uri  # ✅ 여기까지만 redirect_uri
    )

    code = st.query_params.get('code')
    if code:
        # fetch_token에는 redirect_uri 넘기지 않음
        flow.fetch_token(code=code)  # 👈 여기 수정
        creds = flow.credentials
        st.session_state['credentials'] = creds_to_dict(creds)

        st.success("Logged in successfully!")
        st.markdown(
            f'<meta http-equiv="refresh" content="0;url=/" />',
            unsafe_allow_html=True
        )
        st.stop()


def creds_to_dict(creds):
    return {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': creds.scopes
    }

# ===== QR 코드 생성 함수 =====
def generate_qr_code(data):
    qr = qrcode.make(data)
    buffer = io.BytesIO()
    qr.save(buffer, format="PNG")
    return buffer.getvalue()

# ===== 이메일 보내기 함수 =====
def send_email(service, to, subject, body, qr_img_bytes):
    message = MIMEMultipart('related')
    message['to'] = to
    message['subject'] = subject

    # 본문에 QR 코드 삽입
    html_content = f"""
    <html>
    <body>
        {body}<br><br>
        <img src="cid:qr_code">
    </body>
    </html>
    """
    html_part = MIMEText(html_content, 'html')
    message.attach(html_part)

    # QR 코드 이미지 첨부
    image_part = MIMEImage(qr_img_bytes, name='qr_code.png')
    image_part.add_header('Content-ID', '<qr_code>')
    message.attach(image_part)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    send_message = {'raw': raw}
    service.users().messages().send(userId="me", body=send_message).execute()

# ===== Streamlit 앱 =====
st.title("Jiwoong Event QR Code Emailer for TKS")

# 🚨 'code' 있으면 save_credentials(), 없으면 로그인 버튼 보여줌
if 'credentials' not in st.session_state:
    code = st.query_params.get('code')

    if code:
        save_credentials()
        st.stop()
    else:
        oauth_info = st.secrets["oauth"]
        redirect_uri = st.secrets["redirect_uri"]

        flow = Flow.from_client_config(
            {
                "installed": {
                    "client_id": oauth_info["client_id"],
                    "project_id": oauth_info["project_id"],
                    "auth_uri": oauth_info["auth_uri"],
                    "token_uri": oauth_info["token_uri"],
                    "auth_provider_x509_cert_url": oauth_info["auth_provider_x509_cert_url"],
                    "client_secret": oauth_info["client_secret"],
                    "redirect_uris": oauth_info["redirect_uris"]
                }
            },
            scopes=['https://www.googleapis.com/auth/gmail.send'],
            redirect_uri=redirect_uri
        )

        auth_url, _ = flow.authorization_url(prompt='consent')
        st.write(f"Please [login here]({auth_url}) to authorize the app.")
        st.stop()

# Streamlit Secrets에서 service account 정보 읽어오기
service_account_info = st.secrets["service_account"]

# gspread 인증
gc = gspread.service_account_from_dict(service_account_info)

# 🚀 로그인 완료 후 메인 폼
spreadsheet_url = st.text_input("Google Spreadsheet Link")
email_subject = st.text_input("Email Subject")
email_body_template = st.text_area("Email Body (Use {{First Name}})")

if st.button("Generate QR Codes and Send Emails"):
    if spreadsheet_url and email_subject and email_body_template:
        try:
            sheet = gc.open_by_url(spreadsheet_url)
            worksheet = sheet.sheet1
            data = worksheet.get_all_records()
            df = pd.DataFrame(data)

            # 컬럼 존재 체크
            required_cols = {'Email', 'First Name'}
            if not required_cols.issubset(df.columns):
                st.error(f"Spreadsheet must contain the following columns: {', '.join(required_cols)}")
                st.stop()

            gmail_service = get_gmail_service()

            # QR 코드 및 이메일 전송
            for idx, row in df.iterrows():
                uuid = f"{row['Email']}-{row['First Name']}"
                qr_img_bytes = generate_qr_code(uuid)

                email_body = email_body_template.replace('{{First Name}}', row['First Name'])

                send_email(
                    service=gmail_service,
                    to=row['Email'],
                    subject=email_subject,
                    body=email_body,
                    qr_img_bytes=qr_img_bytes
                )

            st.success("All emails sent successfully!")
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Please fill in all the fields.")
