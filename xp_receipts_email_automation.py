import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os  # To access environment variables

# SMTP configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587
your_email = ""
your_password = ""

df = pd.read_csv('xp_receipt_data.csv')

cc_email_list = ["gladwin.delrosario.organizations@gmail.com", "bondenie@gmail.com"]

def send_email(recipient_email, recipient_name, activity_name, activity_category, xp_points):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"XP Receipt for Completing the {activity_name}! 🎉"
    msg['From'] = your_email  # Use your actual email
    msg['To'] = recipient_email
    msg['Cc'] = ', '.join(cc_email_list)  # Add CC addresses statically

    # HTML content for the email
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Microsoft Branding Template</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f4f4f4;
                color: #333333;
                margin: 0;
                padding: 0;
            }}
            .container {{
                max-width: 800px;
                margin: 0 auto;
                background-color: #ffffff;
                padding: 25px;
            }}
            .container1 {{
                float: left;
                margin-right: 50px;
            }}
            .logo {{
                width: 200px;
            }}
            h1 {{
                color: #0078d4;
                margin-top: 20px;
            }}
            p {{
                line-height: 1.5; /* Set line spacing */
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <img src="https://raw.githubusercontent.com/gfdelrosario12/msc-assets/main/header.png" alt="Student Ambassador" style="width: 100%;">
            </header>

            <div>
                <h1>Congratulations, {recipient_name}! 🎉</h1>

                <p>I am thrilled to inform you that you have successfully completed the <strong>{activity_name}</strong>, a/an <strong>{activity_category}</strong>, earning you <strong>{xp_points}</strong> points! 🏆</p>

                <p>These points will contribute to your overall progress, helping you level up within our department once you reach the required thresholds in our membership leveling system.</p>

                <p>Keep up the excellent work in the realm of <strong>Cloud Computing</strong>, your dedication and efforts are paving the way for future success! 🚀</p>

                <p>Your consistent growth will surely pay off as you continue to develop your skills and knowledge. 🎓</p>

                <p>If you have any questions or need further assistance, feel free to reach out. We're here to support your journey every step of the way! 💪</p>

                <p>See you aboard! ✨</p>

                <p>Best regards,</p>
            </div>

            <footer class="footer">
                <div class="container2">
                    <strong>Gladwin Ferdz I. Del Rosario</strong><br>
                    Director<br>
                    Cloud Computing Department<br>
                    PUPM Microsoft Student Community<br>
                    <a href="mailto:gladwin.delrosario.organizations@gmail.com">gladwin.delrosario.organizations@gmail.com</a><br><br>
                </div>
            </footer>
        </div>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_content, 'html'))

    # Only include the recipient and CC in the sending addresses
    to_addresses = [recipient_email] + cc_email_list

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(your_email, your_password)
            server.sendmail(your_email, to_addresses, msg.as_string())
        print(f"Email sent successfully to {recipient_name} and CC'd to {', '.join(cc_email_list)}!")
    except Exception as e:
        print(f"Failed to send email to {recipient_name}: {e}")

# Send emails to all contacts
for index, row in df.iterrows():
    send_email(row['Email'], row['First_Name'], row['Activity_Name'], row['Activity_Category'], row['XP_Points'])
