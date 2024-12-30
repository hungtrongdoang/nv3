import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Thông tin email gửi
sender_email = "hoatrodunn@gmail.com"
sender_password = "123581321345589Dzung"  # Nếu dùng mật khẩu ứng dụng, điền mật khẩu này
receiver_email = "hungtrongdoang@gmail.com"

# Tạo nội dung email
subject = "Test Email from Python"
body = "This is a test email sent from a Python script!"

# Cấu hình email
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject

# Thêm nội dung vào email
message.attach(MIMEText(body, "plain"))

try:
    # Kết nối đến server SMTP của Gmail
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()  # Bắt đầu mã hóa TLS
        server.login(sender_email, sender_password)  # Đăng nhập
        server.sendmail(sender_email, receiver_email, message.as_string())  # Gửi email
        print("Email đã được gửi thành công!")
except Exception as e:
    print(f"Không thể gửi email. Lỗi: {e}")
