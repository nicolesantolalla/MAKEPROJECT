import smtplib #module for sending email
import sched, time   #Sched module to schedule the program to run every certani time
import xlrd
import datetime
import config #module we created with login information. 
import math

def send_email(subject, msg):
    try:
        server = smtplib.SMTP('smtp.gmail.com:587') #If you want to use a different email, need to find the code for that one.
        server.ehlo()
        server.starttls()
        server.login(config.EMAIL_ADDRESS, config.PASSWORD)  #config is from module we created. You can add any sender email, but need to allow less secure apps.
        message = "Subject: {}\n\n{}".format(subject, msg)  #Use this configuration so the Subject appears on the "Subject" line.
        server.sendmail(config.EMAIL_ADDRESS, config.RECEIVER_ADDRESS, message) #need to add RECEIVER_EMAIL to config module.
        server.quit()
        print("Success Email sent!")
    except:
        print("Email failed to send.")

subject = " Assignments Due"
msg = "Hello"

schedule.every().monday.at("08:00").do(send_email, subject, msg) #You can put any day and time, and module will work.
schedule.every().tuesday.at("08:00").do(send_email, subject, msg) #We want the program to send an email every day in the morning, so we added every day.
schedule.every().wednesday.at("08:00").do(send_email, subject, msg)
schedule.every().thursday.at("08:00").do(send_email, subject, msg)
schedule.every().friday.at("08:00").do(send_email, subject, msg)
schedule.every().saturday.at("08:00").do(send_email, subject, msg)
schedule.every().sunday.at("08:00").do(send_email, subject, msg)


while True:  #Program need to be constantly running. But need to add sleep because computer can crash.

    schedule.run_pending() #imported from module "schedule"
    time.sleep(2)  #Module will run every 2 seconds. Can put a longer time if you want the module to send you email only once a day.


#if localtime.startswith("Tue"):
    #send_email(subject, msg)

#localtime= localtime+604800

#df = pd.read_excel ('tasks.xlsx')
#print(df)


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

subject = "Things to do today"
msg="hello"

def excel_date(date1):
    temp = datetime.datetime(1899, 12, 30)
    delta = date1 - temp
    return(math.floor((delta.days) + (float(delta.seconds) / 86400)))

file_location= "/Users/Daniela/PycharmProjects/untitled/tasks.xlsx"
workbook= xlrd.open_workbook(file_location)
sheet=workbook.sheet_by_index(0)

now = datetime.datetime.now()
now_date=(now.strftime("%Y-%m-%d %H:%m"))
today = datetime.date.today()

for i in range(sheet.nrows):
    if sheet.cell_value(i, 1)==int(excel_date(datetime.datetime.now())):
        msg=str(sheet.cell_value(i, 0))
        print(msg)
        send_email(subject, msg)
    else:
        print('not today')
