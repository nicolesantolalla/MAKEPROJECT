import smtplib
import config
import time
import schedule
import xlrd
import datetime
import math


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
msg =""

def excel_date(date1):
    temp = datetime.datetime(1899, 12, 30)
    delta = date1 - temp
    return(math.floor((delta.days) + (float(delta.seconds) / 86400)))

workbook= xlrd.open_workbook("tasks.xlsx")
sheet=workbook.sheet_by_index(0)

now = datetime.datetime.now()

def check_for_assignment():
    List=[]
    for i in range(sheet.nrows):
        if sheet.cell_value(i, 1)==int(excel_date(now)):
            List.append(str(sheet.cell_value(i, 0)))
            # we made a list of the tasks so we could display more than one task due the same day
            msg="Your tasks for today are: \n" + '\n'.join(map(str, List))
            #this puts an enter (\n) between each task
        else:
            print('not today') # this is only to check the console is performing the for loop at the time scheudled
    send_email(subject, msg)
schedule.every().day.at("00:33").do(check_for_assignment)


while True:

    schedule.run_pending()
    time.sleep(10)
