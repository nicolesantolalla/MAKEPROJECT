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

file_location= "/Users/nicolesantolalla/PycharmProjects/MAKE1/tasks (1).xlsx"
# need to copy path of file
workbook= xlrd.open_workbook(file_location)
sheet=workbook.sheet_by_index(0)

#when = schedule.every().day.at("21:24").do(send_email, subject, msg)
#send_time = datetime.time(21, 52, 00)


now = datetime.datetime.now()

List=[]
for i in range(sheet.nrows):
    if sheet.cell_value(i, 1)==int(excel_date(now)):
        List.append(str(sheet.cell_value(i, 0)))
        # we made a list of the tasks so we could display more than one task due the same day
        msg="Your tasks for today are: \n" + '\n'.join(map(str, List))
        #this puts an enter (\n) between each task
    else:
        print('not today')
schedule.every().day.at("00:01").do(send_email, subject, msg)



while True:

    schedule.run_pending()
    time.sleep(10)
