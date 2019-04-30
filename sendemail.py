import smtplib

import config

def send_email(subject, msg):
    try:
        server = smtplib.SMTP('smtp.gmail.com:587')
        server.ehlo()
        server.starttls()
        server.login(config.EMAIL_ADDRESS, config.PASSWORD)
        message = "Subject: {}\n\n{}".format(subject, msg)
        server.sendmail(config.EMAIL_ADDRESS, config.RECEIVER_ADDRESS, message)
        server.quit()
        print("Success Email sent!")
    except:
        print("Email failed to send.")

subject = " Test"
msg = "Hello"

schedule.every().tuesday.at("17:55").do(send_email, subject, msg)

# You can out anytime inside the parehtheisis after "at".

while True:

    schedule.run_pending()
    time.sleep(50)

