import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import json

# ------------------------------
# STEP 1: LOAD INPUT DATA
# ------------------------------
df = pd.read_excel("raw_data.xlsx")
df["Date"] = pd.to_datetime(df["Date"])
df["Total"] = df["Qty"] * df["Amount"]

# ------------------------------
# STEP 2: SUMMARY CALCULATION
# ------------------------------
total_revenue = df["Total"].sum()
top_branch = df.groupby("Branch")["Total"].sum().idxmax()
top_product = df.groupby("Product")["Total"].sum().idxmax()
branch_data = df.groupby("Branch")["Total"].sum()

# ------------------------------
# STEP 3: CREATE EXCEL REPORT
# ------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Clean_Report"
headers = ["Date", "Branch", "Product", "Qty", "Amount", "Total"]
ws.append(headers)
for col in ws[1]:
    col.font = Font(bold=True)
    col.alignment = Alignment(horizontal="center")
for _, row in df.iterrows():
    ws.append(list(row.values))
wb.save("Business_Report.xlsx")

# ------------------------------
# STEP 4: CREATE DASHBOARD CHART IMAGE
# ------------------------------
import matplotlib.pyplot as plt

plt.figure(figsize=(8,5))
bars = plt.bar(branch_data.index, branch_data.values, color='darkblue')  # Dark blue bars
plt.title("Branch Performance")
plt.xlabel("Branch")
plt.ylabel("Revenue")


# Add data labels on top of bars
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, height + 0.05*branch_data.max(),
             f'{int(height)}', ha='center', va='bottom', color='black', fontweight='bold')

plt.tight_layout()
plt.savefig("Dashboard.png")
plt.close()


# ------------------------------
# STEP 5: SEND EMAIL WITH ATTACHMENTS AND IMAGE
# ------------------------------
config = json.load(open("config.json"))
sender = config["sender"]
password = config["password"]
receiver = config["receiver"]

# Email setup
msg = MIMEMultipart()
msg["Subject"] = "Daily Business Report - Automation"
msg["From"] = sender
msg["To"] = receiver

# Email body
email_body = f"""
Hello Team,

Here is the Daily Business Summary:

Total Revenue : {total_revenue}
Top Branch    : {top_branch}
Top Product   : {top_product}

Branch-wise Revenue:
{branch_data.to_string()}

Reports Generated:
1. Business_Report.xlsx
2. Dashboard.png

Regards,
Automation Ganesh
"""
msg.attach(MIMEText(email_body, "plain"))

# Attach Excel report
with open("Business_Report.xlsx", "rb") as f:
    attach_file = MIMEApplication(f.read(), _subtype="xlsx")
    attach_file.add_header('Content-Disposition','attachment',filename="Business_Report.xlsx")
    msg.attach(attach_file)

# Attach dashboard image
with open("Dashboard.png", "rb") as f:
    img = MIMEImage(f.read())
    img.add_header("Content-Disposition", "attachment", filename="Dashboard.png")
    msg.attach(img)

# Send email
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender, password)
server.send_message(msg)
server.quit()

print("Email with report and dashboard image sent successfully!")
