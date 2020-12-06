from ftplib import FTP
from datetime import datetime
import os
import xlrd
import time
import re
import sys


while True:
    HMI_list = ['10.3.1.112', '10.3.1.125', '10.3.1.203', '10.3.1.205', '10.3.1.207', '10.3.1.209', '10.3.1.211',
                '10.3.1.213', '10.3.1.215', '10.3.1.217', '10.3.1.219', '10.3.1.221', '10.3.1.223', '10.3.1.225',
                '10.3.1.227']
    PODs = ['P0S1', 'P0S2', 'P1S1', 'P2S1', 'P2S2', 'P3S1', 'P3S2', 'P4S1', 'P4S2', 'P5S1', 'P5S2', 'P6S1', 'P6S2', 'P7S1',
            'P7S2']

    date = input('Please enter date in format yyyymmdd:')  # input date

    if not re.match("^[0-9]*$", date):
        print("Error! Only numbers allowed!")
        time.sleep(2)
        sys.exit()
    elif len(date) != 8:
        print("Error! Only 8 characters allowed!")
        time.sleep(2)
        sys.exit()
    start = datetime.now()
    filename = date + '.dtl'
    averagefortheday = []  # final result array
    GMaverage = []  # GM average fans

    # Making new directory

    os.getcwd()
    savedirectory = os.getcwd() + '\\'
    # define the name of the directory to be created
    pathfornewdirectory = os.getcwd() + '\\' + 'temp'
    try:
        os.mkdir(pathfornewdirectory)
    except OSError:
        print("Creation of the directory %s failed" % pathfornewdirectory)
        print("Directory already exists, downloading files...")
    else:
        print("Successfully created the directory %s " % pathfornewdirectory)
        print("Downloading files...")
    os.chdir(pathfornewdirectory)
    savedirectorytemp = os.getcwd() + '\\'


    def getFile():
        ftp = FTP(HMI)
        path = '/datalog/Fan Speeds/'
        ftp.login(user='uploadhis', passwd='111111')
        ftp.encoding = "utf-8"
        ftp.cwd(path)
        ftp.retrbinary("RETR " + filename, open(savedirectorytemp + filename, 'wb').write, 1024)
        ftp.quit()


    def getaveragefromexcel():
        excelfilepath = str(savedirectorytemp + date + '.xls')
        time.sleep(0.01)
        wb = xlrd.open_workbook(excelfilepath)
        sheet = wb.sheet_by_index(0)
        arraydata = []
        GMarraydata = []
        if HMI != '10.3.1.203':
            for j in range(sheet.nrows - 1):
                for i in range(sheet.ncols - 3):
                    data = (sheet.cell_value((1 + j), (3 + i)))
                    if data != 0:
                        arraydata.append(data)
        else:
            for j in range(sheet.nrows - 1):
                for i in range(sheet.ncols - 8):
                    data = (sheet.cell_value((1 + j), (3 + i)))
                    if data != 0:
                        arraydata.append(data)
            for j in range(sheet.nrows - 1):
                for i in range(sheet.ncols - 9):
                    GMdata = (sheet.cell_value((1 + j), (9 + i)))
                    if GMdata != 0:
                        GMarraydata.append(GMdata)
            time.sleep(0.01)
            GMaverage.append(sum(GMarraydata) / len(GMarraydata))
        average = sum(arraydata) / len(arraydata)
        averagefortheday.append(average)


    def deletefilesintemp():
        os.remove(filename)
        os.remove(date + '.xls')



    for HMI in HMI_list:
        getFile()  # download .dtl file from ftp
        time.sleep(0.01)
        os.startfile(filename)  # convertfile from .dtl to .xls
        time.sleep(0.1)
        getaveragefromexcel()  # calculate average from .xls
        time.sleep(0.01)
        deletefilesintemp()  # deletes files .dtl and .xls
        time.sleep(0.01)
    round_averagefortheday = [round(num, 2) for num in averagefortheday]

    time.sleep(0.01)
    print('Results for '+ date[0:4] + '/'+ date[4:6] + '/' + date[6:8] + ':')
    for n in range(len(PODs)):
        print(PODs[n] + ' = ' + str(round_averagefortheday[n]))
    print('GM = ' + str(round(GMaverage[0], 2)))
    print(*PODs)
    print(*round_averagefortheday)
    end = datetime.now()
    diff = end - start
    milisecondsdiff = "{:.2f}".format(diff.seconds)
    print('Average calculated in ' + milisecondsdiff + 's')

    while True:
        answer = str(input('Run again? (y/n): '))
        if answer in ('y', 'n'):
            break
        print("invalid input.")
    if answer == 'y':
        os.chdir(savedirectory)
        continue
    else:
        print("Goodbye")
        break
