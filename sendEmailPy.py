#!/usr/bin/python
__author__ = 'pgalindo'

import os
import socket
import smtplib
import base64
import xlrd
import csv , operator
import shutil, os
import time
import glob
import getpass
from os import walk
from xlsxwriter.workbook import Workbook
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#funtion to list files in a folder
def ls(ruta = '.'):
    dir, subdirs, archivos = next(walk(ruta))
    print(archivos)

username = input("Enter email account: ")
password = input("Enter password: ")
actualPath = os.getcwd() #to obtain path
print("Files of this folder: ")
ls()
fileName = input("Enter the file name (.xlsx) to start the automatic process(IMPORTANT not include file extension): ")
pathToFileName = actualPath + os.sep + fileName + '.xlsx'

# Function that send an email.
def sendMail(username, password, from_addr, to_addrs, message):
    msg = MIMEMultipart()

    msg['Subject'] = "Testing automating email"
    msg['From'] = from_addr
    msg['To'] = to_addrs

    # Attach HTML to the email
    body = MIMEText('text', "plain")
    body.set_payload(message)
    msg.attach(body)

    # Attaching file if is needed
    att1 = MIMEApplication(open("PathToFile", "rb").read())
    att1.add_header('Content-Disposition', 'attachment', filename="MyFileName.ext")
    msg.attach(att1)

    #att2 = MIMEApplication(open("file2.pdf", "rb").read())
    #att2.add_header('Content-Disposition', 'attachment', filename="file2.pdf")
    #msg.attach(att2)

    server = smtplib.SMTP('smtp.office365.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(username, password)
    server.sendmail(from_addr, to_addrs, msg.as_string())
    print ("Email successfully sent to", to_addrs)
    server.quit()

# Function to convert a .xslx (excel) to a csv file
def csv_from_excel(PathToExcel):
    wb = xlrd.open_workbook(PathToExcel)
    sh = wb.sheet_by_index(0)
    with open('tmpCsvFile.csv', 'w',newline='') as csv_file:
        #wr = csv.writer(csv_file, delimiter=',', quoting=csv.QUOTE_MINIMAL)
        wr = csv.writer(csv_file)

        for rownum in range(sh.nrows):
            wr.writerow(sh.row_values(rownum))
    csv_file.close()

csv_from_excel(pathToFileName)
os.remove(pathToFileName) #remove previous files

# Function to check if is needed to send an email
def checkCsvRows(PathToInitialCsv, PathToFinalCsv):
    with open(PathToFinalCsv, 'w', newline='') as outputCsv:  
        with open(PathToInitialCsv, 'r', encoding='latin-1') as csvfile:
            reader = csv.DictReader(csvfile)
            output = csv.DictWriter(outputCsv, fieldnames=reader.fieldnames)
            output.writeheader()
            for row in reader:
                if len(row['Remediation due date']) == 0 or datetime.strptime(row['Remediation due date'], '%d/%m/%Y') < datetime.strptime(datetime.now().strftime('%d/%m/%Y'), '%d/%m/%Y'):
                    if len(row['Email']) != 0:
                        message = 'The message' + row['field'] + row['field2'] + 'goodbye.'
                        sendMail(username, password, username, row['Email'], message)
                    else:
                        print("Email don't exist for this row: " + row)    
                    if len(row['Counter']) == 0:
                        row['Counter'] = str(0)
                    row['Counter'] = str(float(row['Counter']) + 1)    
                output.writerow(row)

            outputCsv.close()        
        csvfile.close()  

pathInitialCsv = actualPath + os.sep + 'tmpCsvFile.csv'
pathFinalCsv = actualPath + os.sep + fileName + '.csv'
checkCsvRows(pathInitialCsv, pathFinalCsv)
os.remove(pathInitialCsv) #remove previous files

# Function to generate a excel file from a .csv file
def excel_from_csv(PathToFinalCsv):
    for PathToFinalCsv in glob.glob(os.path.join('.','*.csv')):
        workbook = Workbook(PathToFinalCsv[:-4] + '.xlsx')
        worksheet = workbook.add_worksheet()
        with open(PathToFinalCsv, 'rt', encoding='latin-1') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()

excel_from_csv(pathFinalCsv)
os.remove(pathFinalCsv) #remove previous files
