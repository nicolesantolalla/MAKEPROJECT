import smtplib #module for sending email
import sched, time   
import pandas as pd
import xlrd
import datetime
import config

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


while True:  #Program need to be constantly running. But need to add sleep because computer can crash.

    schedule.run_pending() #imported from module "schedule"
    time.sleep(2)  #Module will run every 2 seconds. Can put a longer time if you want the module to send you email only once a day.


localtime=time.asctime(time.localtime(time.time()))
print("the time is " + localtime)

#if localtime.startswith("Tue"):
    #send_email(subject, msg)

#localtime= localtime+604800

#df = pd.read_excel ('tasks.xlsx')
#print(df)

file_location= "/Users/Daniela/PycharmProjects/untitled/tasks.xlsx"
workbook= xlrd.open_workbook(file_location)
sheet=workbook.sheet_by_index(0)
due=sheet.cell_value(9,1)
due_date=xlrd.xldate_as_tuple(due, 0)
due_date2=datetime.datetime(*due_date)
print(due_date2.strftime("%Y-%m-%d"))

now = datetime.datetime.now()
print ("Current date and time using str method of datetime object:")
now_date=(now.strftime("%Y-%m-%d"))
print(now_date)


if due_date2.strftime("%Y-%m-%d")== (now.strftime("%Y-%m-%d")):
    send_email(subject, msg)
