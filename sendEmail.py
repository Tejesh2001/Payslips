# import email, smtplib, ssl

# from email import encoders
# from email.mime.base import MIMEBase
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText

# def sendEmail(name):

#     subject = "Salary slips check"
#     body = "This is an email with attachment sent from Python"
#     sender_email = "hr@conacent.com"
#     receiver_email = "chandrani@conacent.com"
#     password = "conacentcf90"

#     # Create a multipart message and set headers
#     message = MIMEMultipart()
#     message["From"] = sender_email
#     message["To"] = receiver_email
#     message["Subject"] = subject
#     message["Bcc"] = receiver_email  # Recommended for mass emails

#     # Add body to email
#     message.attach(MIMEText(body, "plain"))

#     filename = name  # In same directory as script

#     # Open PDF file in binary mode
#     with open(filename, "rb") as attachment:
#         # Add file as application/octet-stream
#         # Email client can usually download this automatically as attachment
#         part = MIMEBase("application", "octet-stream")
#         part.set_payload(attachment.read())

#     # Encode file in ASCII characters to send by email    
#     encoders.encode_base64(part)

#     # Add header as key/value pair to attachment part
#     part.add_header(
#         "Content-Disposition",
#         f"attachment; filename= {filename}",
#     )

#     # Add attachment to message and convert message to string
#     message.attach(part)
#     text = message.as_string()

#     # Log in to server using secure context and send email
#     context = ssl.create_default_context()
#     with smtplib.SMTP_SSL("mail.conacent.com", 465, context=context) as server:
#         server.login(sender_email, password)
#         server.sendmail(sender_email, receiver_email, text)


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def sendEmail(name, email, month, person_name, type_of_file, is_pan):

    body = f''' Dear  {person_name}, 
                Please find attached {type_of_file} {month}.
                {is_pan}
                HR'''
    # put your email here
    sender = 'hr@conacent.com'
    # get the password in the gmail (manage your google account, click on the avatar on the right)
    # then go to security (right) and app password (center)
    # insert the password and then choose mail and this computer and then generate
    # copy the password generated here
    password = 'ConacentCf90!@'
    # put the email of the receiver here
    receiver = email

    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = receiver
    message["Bcc"] = sender
    message['Subject'] = f' {month}'

    message.attach(MIMEText(body, 'plain'))

    pdfname = name

    # open the file in bynary
    binary_pdf = open(pdfname, 'rb')

    payload = MIMEBase('application', 'octate-stream', Name=pdfname)
    # payload = MIMEBase('application', 'pdf', Name=pdfname)
    payload.set_payload((binary_pdf).read())

    # enconding the binary into base64
    encoders.encode_base64(payload)

    # add header with pdf name
    payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
    message.attach(payload)

    #use gmail with port
    session = smtplib.SMTP('mail.conacent.com', 587)
    #enable security
    session.starttls()

    #login with mail_id and password
    session.login(sender, password)

    text = message.as_string()
    session.sendmail(sender, receiver, text)
    session.quit()
    