import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Load the Excel file
file_path = 'spons.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Extract company names and Gmail addresses
company_data = df[['Company Name', 'Email']]
gmail_data = company_data[company_data['Email'].str.contains('@', na=False)]
company_list = gmail_data.to_dict(orient='records')

# Email credentials
GMAIL_USER = 'your-email@gmail.com'
GMAIL_PASSWORD = 'your-app-specific-password'

# Define the function to send HTML emails with attachments
def send_email_with_attachment(to_email, subject, body, attachment_path):
    # Set up the MIME (multi-part message)
    msg = MIMEMultipart()
    msg['From'] = GMAIL_USER
    msg['To'] = to_email
    msg['Subject'] = subject

    # Attach the email body to the MIME message (HTML version)
    msg.attach(MIMEText(body, 'html'))

    # Check if the attachment exists
    if attachment_path and os.path.isfile(attachment_path):
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            msg.attach(part)
            print(f"Attachment {os.path.basename(attachment_path)} added.")
    else:
        print("No valid attachment found.")

    try:
        # Set up the Gmail SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Start TLS encryption
        server.login(GMAIL_USER, GMAIL_PASSWORD)

        # Send the email
        server.sendmail(GMAIL_USER, to_email, msg.as_string())
        print(f"Email sent to {to_email}")

        server.quit()  # Close the server connection

    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {e}")

# Email subject and body template
subject = "Sponsorship Opportunity for Odyssey'25"
attachment_path = "Odyssey'25-Sponsorship-proposal.pdf"  # Example: 'proposal.pdf'

# Send a personalized email to each company
for company in company_list:
    name = company['Company Name']
    print(f"Processing {company['Company Name']} ({company['Email']})")  # Debugging print

    email = company['Email']

    # Personalize the email body with bold text for the company name
    body = f"""
    <html>
    <body>
        <p>Greetings from Indraprastha Institute of Information Technology, Delhi!</p>

        <p>We are thrilled to introduce <strong>Odyssey'25</strong>, the Annual Cultural Fest hosted by IIIT Delhi, a leading institute in the field of information technology, established by the Government of Delhi in 2008. Odyssey'25 is set to take place on the 17th and 18th of January, 2025, bringing together students, artists, and cultural enthusiasts from across India for an unforgettable celebration.</p>

        <p>Odyssey'25 is a two-day extravaganza featuring a vibrant mix of cultural performances and diverse events, spanning dance, music, art, fashion, theater, and more. The festival promises a stellar lineup of multiple singer performances and comedy shows, which have become crowd favorites over the years. Past editions of Odyssey have seen a remarkable footfall of over 100,000 attendees, including students from across India and people from all walks of life.</p>

        <p>We’ve had the privilege of collaborating with renowned artists such as Shankar Ehsaan Loy, KrSna, and Aakash Gupta, Salim Sulaiman, Jubin Nautiyal, Gajendra Verma, Sam Altman, Rahul Subramaniam, Zakir Khan. With each year, Odyssey continues to grow in scale and reach, becoming a beacon of Delhi’s cultural heritage.</p>

        <p>We extend a warm invitation to <strong>{name}</strong> to partner with us for <strong>Odyssey'25</strong>, offering a unique opportunity to showcase your brand to a diverse and dynamic audience. Last year, our fest drew nearly 1 Lakh attendees, making it a fantastic platform for marketing and business development. For your convenience, we have attached our fest brochure for your perusal. We believe that your association with Odyssey'25 will not only elevate our fest’s grandeur but also provide your company with exceptional visibility.</p>

        <p>Please feel free to reach out for any further information. My contact details are:</p>

        <p>Saksham Bansal<br>
        saksham23467@iiitd.ac.in</p>

        <p>We look forward to a successful and meaningful partnership with <strong>{name}</strong> for Odyssey'25!</p>

        <p>Best Regards,<br>
        Saksham Bansal<br>
        Sponsorship Team<br>
        Odyssey’25 | IIIT Delhi</p>
    </body>
    </html>
    """

    # Send the email with attachment
    send_email_with_attachment(email, subject, body, attachment_path)
