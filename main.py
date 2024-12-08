import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import time

# Load and clean the Excel file
file_path = 'D:/gmailPY/finalEditAcc.xlsx'  #* Update your file path here
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Clean the column names
df.columns = df.columns.str.strip()  # Remove leading/trailing whitespace

# Check if required columns exist
required_columns = ['Name', 'Email', 'Committee']
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"Missing columns in the Excel file: {missing_columns}")
    exit()

# Fix the dataframe structure
df = df[required_columns].dropna()  # Ensure these match the actual column names

# Prompt for email credentials
SENDER_EMAIL = input("Enter your email: ")
SENDER_PASSWORD = input("Enter your app password: ")

# Email configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Set up the SMTP server
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
except smtplib.SMTPAuthenticationError:
    print("Authentication failed. Check your email and password.")
    exit()
except Exception as e:
    print(f"Error setting up SMTP server: {e}")
    exit()

# Declare a counter variable
counter = 0

# Iterate through each row in the Excel sheet
for index, row in df.iterrows():
    recipient_email = row['Email']
    recipient_name = row['Name']
    committee = row['Committee']

    # Updated email body with committee in specified color and bold
    body = f"""
    <html>
      <body style="background-color: #0f212b; color: #ffffff; font-family: Arial, sans-serif; padding: 20px; position: relative;">
        
        <img src="https://drive.google.com/uc?id=12JkGCXpaXnsj5EXUGPrgW_w9pYLbJ-LO" alt="Header Image" style="width: 100%; height: auto; margin-bottom: 20px;">
        
        <p>Dear <strong style="color:#ffc222;"> {recipient_name} </strong>,</p>
        <p>Congratulations! We are thrilled to welcome you to the Enactus Menoufia's <strong style="color:#ffc222;">{committee} committee</strong>. Your skills, passion, and dedication impressed us throughout the interview process, and we’re excited to have you join us in our mission.</p>
        <p>As a member of the <strong style="color:#ffc222;">{committee} committee</strong>, you’ll have the opportunity to make a positive impact, work alongside other driven students, and further develop your abilities. We’re eager to see the unique contributions you’ll bring to the team.</p>
        <p>Please keep an eye out for our upcoming emails with details about orientation, upcoming projects, and your role in our team. Feel free to reach out with any questions or if there’s anything we can assist you with as you transition into your new role.</p>
        <p>Welcome aboard!</p>
        <p>Best regards,<br>Enactus</p>

        <img src="https://enactusegypt.org/wp-content/uploads/2021/01/Enactus-Full-Color-2.png" alt="Logo" style="position: absolute; bottom: 10px; right: 10px; width: 100px; height: auto;">
      </body>
    </html>
    """

    # Create the email content
    message = MIMEMultipart()
    message['From'] = SENDER_EMAIL
    message['To'] = recipient_email
    message['Subject'] = f"Congratulations and Welcome to the {committee} Committee at Enactus Menoufia!"

    # Attach the body as HTML
    message.attach(MIMEText(body, 'html'))

    # Send the email
    try:
        server.sendmail(SENDER_EMAIL, recipient_email, message.as_string())
        counter += 1
        print(f"Email sent successfully to {recipient_name} ({recipient_email}). (Total sent: {counter})")
        time.sleep(1)  # Add a delay to prevent email throttling
    except Exception as e:
        print(f"Failed to send email to {recipient_name} ({recipient_email}): {e}")

# Close the SMTP server
server.quit()
print(f"Total emails sent: {counter}")
