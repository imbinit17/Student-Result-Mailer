import smtplib
from email.message import EmailMessage

def sendMail(Receiver_Email,Subject,message,files):
    Sender_Email = "example@gmail.com" # PLEASE ENTER YOUR EMAIL ADRESS
    Password = "password" # HERE YOUR PASSWORD

    newMessage = EmailMessage()                         
    newMessage['Subject'] = Subject
    newMessage['From'] = Sender_Email                   
    newMessage['To'] = Receiver_Email                   
    newMessage.set_content(message)

    
    for file in files:
        with open(file, 'rb') as f:
            file_data = f.read()
            file_name = f.name
        newMessage.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        
        smtp.login(Sender_Email, Password)              
        smtp.send_message(newMessage)
