import pymysql
import datetime
import time
import unidecode
from unidecode import unidecode
import os
from random import randint
import reportlab

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

import connection
from connection import db_ip_address
from connection import db_user
from connection import db_password
from connection import db_name
from connection import email_sender
from connection import email_sender_password
from connection import email_hiwi
from connection import server_out
from connection import port_out
from connection import receipt_tutor_folder


def random_with_N_digits(n):
    range_start = 10**(n-1)
    range_end = (10**n)-1
    return randint(range_start, range_end)



def sendReceipt(fname, lname, event_description, receipt_path, receipt_name, email_tutor, receipt_type):
    global email_sender
    global email_sender_password
    global server_out
    global port_out


    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_tutor

    if receipt_type == 1:
        msg['Subject'] = "TUMi: " + event_description + " Money Collection Receipt"
    elif receipt_type == 2:
        msg['Subject'] = "TUMi: " + event_description + " Office Transaction"

    body = "Hi " + fname + " " + lname + ",\n\nAttached herewith is the receipt from our recent transaction.\n\n\nRegards,\nThomas Bergmann, M.A.\nOffice for Incoming Exchange Students / TUMi\n\nTechnical University of Munich\nTUM International Centre\n\nArcisstr. 21\n80333 Munich\n\nTel: + 49 89 289 25477"
    msg.attach(MIMEText(body, 'plain'))
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(receipt_path, 'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + receipt_name)
    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP(server_out, port_out)
    server.starttls()
    server.login(email_sender, email_sender_password)
    
    #email_receivers = [email_tutor] + [email_sender] + [email_hiwi]
    email_receivers = [email_tutor]
    server.sendmail(email_sender, email_receivers, text)
    
    print("\nReceipt sent to:", email_tutor)
    server.quit()


def create_Cash_Collection_Receipt(event_description, event_code, event_date, fname, lname, money_out, email_tutor):
    global receipt_tutor_folder

    receipt_number = str(random_with_N_digits(8))
    receipt_name = receipt_number + "_" + unidecode(fname + lname) + ".pdf"
    receipt_path = receipt_tutor_folder + receipt_name
    receipt_event_description = "Event name                             : " + event_description
    receipt_event_code = "Event code                              : " + event_code
    receipt_event_date = str(event_date)
    year, month, day = map(str, receipt_event_date.split('-'))
    receipt_event_date = day + "/" + month + "/" + year
    receipt_event_date = "Event date                               : " + str(receipt_event_date)

    name = "Tutor                                        : " + fname + " " + lname
    system_timestamp = time.strftime('%d/%m/%Y  %X')
    receipt_timestamp = "Date and time of collection      :" + " " + system_timestamp

    amount = "Amount collected                     : €" + money_out
    receipt_number = "Receipt number                       : " + receipt_number
    save_name = os.path.join(os.path.expanduser("~"), receipt_tutor_folder, receipt_name)
    c = canvas.Canvas(save_name, pagesize=A4)
    # header text
    c.setFont(size=14, psfontname="Helvetica", leading=None)
    c.drawString(70, 670, "International Centre")
    c.drawString(70, 650, "Technical University of Munich")
    c.drawString(70, 630, "Arcisstr. 21, 80333 Munich")
    c.drawString(70, 610, "Germany")

    c.setFont(size=22, psfontname="Helvetica-Bold", leading=None)
    c.drawCentredString(x=300, y=570, text="RECEIPT")

    c.setFont(size=14, psfontname="Helvetica", leading=None)
    c.drawString(70, 530, receipt_event_description)
    c.drawString(70, 510, receipt_event_code)
    c.drawString(70, 490, receipt_event_date)

    c.drawString(70, 450, name)
    c.drawString(70, 430, receipt_timestamp)
    c.drawString(70, 410, amount)

    c.drawString(70, 370, receipt_number)

    c.drawImage(image='index.png', x=295, y=700, preserveAspectRatio=False, width=229.2, height=79.2)
    c.showPage()
    c.save()

    sendReceipt(fname, lname, event_description, receipt_path, receipt_name, email_tutor, receipt_type=1)



def create_Historical_Receipt(description, event_date, fname, lname, money_flow, email_tutor):
    global receipt_tutor_folder


    receipt_number = str(random_with_N_digits(8))
    receipt_name = receipt_number + "_" + unidecode(fname + lname) + ".pdf"
    receipt_path = receipt_tutor_folder + receipt_name
    receipt_description = "Description                              : " + description
    event_date = str(event_date)
    year, month, day = map(str, event_date.split('-'))
    event_date = day + "/" + month + "/" + year
    event_date = "Event date                               : " + str(event_date)

    name = "Tutor                                        : " + fname + " " + lname
    system_timestamp = time.strftime('%d/%m/%Y  %X')
    receipt_timestamp = "Date and time of transaction    :" + " " + system_timestamp

    amount = "TUMi cash flow                        : €" + money_flow
    receipt_number = "Receipt number                       : " + receipt_number
    save_name = os.path.join(os.path.expanduser("~"), receipt_tutor_folder, receipt_name)
    c = canvas.Canvas(save_name, pagesize=A4)
    # header text
    c.setFont(size=14, psfontname="Helvetica", leading=None)
    c.drawString(70, 670, "International Centre")
    c.drawString(70, 650, "Technical University of Munich")
    c.drawString(70, 630, "Arcisstr. 21, 80333 Munich")
    c.drawString(70, 610, "Germany")

    c.setFont(size=22, psfontname="Helvetica-Bold", leading=None)
    c.drawCentredString(x=300, y=570, text="RECEIPT")

    c.setFont(size=14, psfontname="Helvetica", leading=None)

    c.drawString(70, 510, receipt_description)
    c.drawString(70, 490, event_date)

    c.drawString(70, 450, name)
    c.drawString(70, 430, receipt_timestamp)
    c.drawString(70, 410, amount)

    c.drawString(70, 370, receipt_number)

    c.drawImage(image='index.png', x=295, y=700, preserveAspectRatio=False, width=229.2, height=79.2)
    c.showPage()
    c.save()

    sendReceipt(fname, lname, description, receipt_path, receipt_name, email_tutor, receipt_type=2)




# Main Body


db = pymysql.connect(host=db_ip_address, user=db_user, passwd=db_password, db=db_name)
cursor = db.cursor()

option = input("\nWhat would you like to do?\n(1) Add entry\t\t\t(3) Current cash\t\t\t(5) Time range balance\n(2) Tutor cash out\t\t(4) Display transactions\t(6) Historical event money flow")

if option == '1':
    description = input("Enter description:")
    i = 0
    while i == 0:
        try:
            sum = input("Money flow:")
            sum = float(sum)
            i = 1
        except:
            print("Please enter a number with max 2 decimal places.")

    sql = 'INSERT INTO cashflow value(now(), "%s", "%.2f")' % (description, sum)
    cursor.execute(sql)
    db.commit()


elif option == '2':
    i = 0
    while i == 0:
        try:
            email_tutor = input("Email of tutor collecting the money:")
            sql = 'SELECT * FROM tutor_data WHERE Email = "%s"' % (email_tutor)
            cursor.execute(sql)
            result = cursor.fetchall()
            for column in result:
                fname = column[0]
                lname = column[1]
            print(fname, lname, "is in the tutor list.\n")
            i = 1
        except:
            print("The person is not in the tutor list.\n")

    i = 0
    while i == 0:
        try:
            event_code = input("Event Code:")
            sql1 = 'SELECT * FROM events WHERE eventCode = "%s"' % (event_code)
            cursor.execute(sql1)
            result2 = cursor.fetchall()
            for column in result2:
                event_date = column[0]
                event_description = column[-1]  # For receipt
            entry_description = event_description + " tutor cash collection"  # For databank. Variable description to check if event code exists. Otherwise error

            print("Event", event_code, "selected.\n")

            j = 0
            while j == 0:
                try:
                    cash_out = input("Money flow:")
                    cash_out = float(cash_out)
                    sql1 = 'INSERT INTO cashflow value(now(), "%s", "%.2f")' % (entry_description, cash_out)
                    cursor.execute(sql1)
                    money_out = '%.2f' % (abs(cash_out))
                    create_Cash_Collection_Receipt(event_description, event_code, event_date, fname, lname, str(money_out), email_tutor)
                    db.commit()
                    j = 1
                    i = 1
                except:
                    print("Please enter a number with max 2 decimal places.")
        except:
            print("Wrong event code. Please try again.\n")


elif option == '3':
    sql = "SELECT sum(Cash_Flow) FROM cashFlow"
    cursor.execute(sql)
    result = cursor.fetchall()
    for column in result:
        print("\nCurrent total cash in office =", column[0], "Euros.")


elif option == '4':
    i = 0
    while i == 0:
        try:
            start_date_BrE = input('\nStart of the date range in DD/MM/YYYY format:')  # British English format
            day, month, year = map(int, start_date_BrE.split('/'))
            start_date = datetime.datetime(year, month, day, hour=0, minute=0, second=0)  # Correct SQL date format

            end_date_BrE = input('End of the date range in DD/MM/YYYY format:')  # British English format
            day2, month2, year2 = map(int, end_date_BrE.split('/'))
            end_date = datetime.datetime(year2, month2, day2, hour=23, minute=59, second=59)  # Correct SQL date format

            sql = "SELECT * FROM cashFlow WHERE Time_Stamp BETWEEN '%s' AND '%s'" % (start_date, end_date)
            cursor.execute(sql)
            result = cursor.fetchall()
            for column in result:
                print(column[0], "\t", column[1], "\t\t\t\t", column[2])

            i = 1

        except:
            print("Please enter the date in the correct format.\n")
            continue


elif option == '5':
    i = 0
    while i == 0:
        try:
            start_date_BrE = input('\nStart of the date range in DD/MM/YYYY format:')  # British English format
            day, month, year = map(int, start_date_BrE.split('/'))
            start_date = datetime.datetime(year, month, day, hour=0, minute=0, second=0)  # Correct SQL date format

            end_date_BrE = input('End of the date range in DD/MM/YYYY format:')  # British English format
            day2, month2, year2 = map(int, end_date_BrE.split('/'))
            end_date = datetime.datetime(year2, month2, day2, hour=23, minute=59, second=59)  # Correct SQL date format

            sql5 = 'SELECT * FROM cashflow WHERE Time_Stamp BETWEEN "%s" AND "%s"' % (start_date, end_date)
            cursor.execute(sql5)
            result = cursor.fetchall()
            amount = 0
            for column in result:
                amount += column[2]

            print("\nTotal cash in office between", start_date_BrE, "and", end_date_BrE, "=", amount, "Euros")
            i = 1

        except:
            print("Please enter the date in the correct format.\n")
            continue


elif option == '6':
    i = 0
    while i == 0:
        try:
            email_tutor = input("Email of tutor: ")
            sql = 'SELECT * FROM tutor_data WHERE Email = "%s"' % (email_tutor)
            cursor.execute(sql)
            result = cursor.fetchall()
            for column in result:
                fname = column[0]
                lname = column[1]
            print(fname, lname, "is in the tutor list.\n")
            i = 1
        except:
            print("The person is not in the tutor list.\n")

    i = 0
    while i == 0:
        try:
            description = input("Description:")
            event_date = input('Date of event in DD/MM/YYYY format')
            day, month, year = map(int, event_date.split('/'))
            testDate = datetime.date(year, month, day)

            j = 0
            while j == 0:
                try:
                    money_flow = input("Money flow:")
                    money_flow = float(money_flow)
                    sql1 = 'INSERT INTO cashflow value(now(), "%s", "%.2f")' % (description, money_flow)
                    cursor.execute(sql1)
                    money_out = '%.2f' % (money_flow)
                    create_Historical_Receipt(description, testDate, fname, lname, str(money_flow), email_tutor)

                    db.commit()
                    j = 1
                    i = 1
                except:
                    print("Please enter a number with max 2 decimal places.")
        except:
            print("Please enter the date in correct format.\n")


else:
    print("\nPlease restart and enter a digit from 1 - 6 next time.")

db.close()




