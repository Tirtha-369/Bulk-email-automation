import smtplib
import json
import pandas as pd
import time
from email.message import EmailMessage

# -------- LOAD CONFIG --------
with open("config.json") as f:
    config = json.load(f)

# -------- READ EXCEL --------
df = pd.read_excel("hr1_contacts.xlsx")

# -------- LOGIN --------
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(config["sender_email"], config["app_password"])

print("Login successful")

# -------- LOOP THROUGH EXCEL --------
for index, row in df.iterrows():

    receiver_name = row["name"]
    receiver_email = row["email"]

    msg = EmailMessage()
    msg["From"] = config["sender_email"]
    msg["To"] = receiver_email
    msg["Subject"] = config["subject"]

    # Personalize body
    personalized_body = f"Hi {receiver_name},\n\n" + config["body"]

    msg.set_content(personalized_body)

    # Attach files
    for file_path in config["attachments"]:
        with open(file_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename=file_path.split("/")[-1]
            )

    server.send_message(msg)
    print(f"Email sent to {receiver_name} ({receiver_email})")

    time.sleep(5)  # prevent spam detection

server.quit()
print("All emails sent successfully")
