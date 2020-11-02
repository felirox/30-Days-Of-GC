import requests
import openpyxl as excel
import time
from datetime import date as dtt
import xlwt 
from xlwt import Workbook 
  
# Workbook is created 
wb = Workbook() 

sheet1 = wb.add_sheet('Sheet 1')

def readContacts(fileName):
    
    lst = []
    i = 1
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    secCol = sheet['B']
    fifCol = sheet['E']
    sixCol = sheet['F']
    sevCol = sheet['G']
    eigCol = sheet['H']
    tdict=9999999999
    cell = 1
    hoursRemaining = 99
    minRemaining = 99
    for cell in range(len(firstCol)):
        stu_name = (firstCol[cell].value)
        stu_mail = (secCol[cell].value)
        skbadge_1 = (fifCol[cell].value)
        skbadgename_1 = (sixCol[cell].value)
        skbadge_2 = (sevCol[cell].value)
        skbadgename_2 = (eigCol[cell].value)
        opbad=""
        opgood=""
        if(int(skbadge_1)==0):
            opbad="Please Hurry up!! You have not completed any badge yet. Please start doing the same soon. Approach the community if you need help! "
        if(int(skbadge_1)==1):
            opbad="Hurry up!! You have completed only one skill badge. Good progress! Please start doing the other labs soon."
        if(int(skbadge_1)==2):
            opgood="Hurry up!! You have completed only two out of six badges. Pick up your pace and get this bread!"
        if((int(skbadge_1)>2 and int(skbadge_1)<6) or (int(skbadge_2)>2 and int(skbadge_2)<6) ):
            opgood="Great work! You are almost there, but hurry up! Keep up with your pace and complete the labs on time!!"
        if((int(skbadge_1)==6) or (int(skbadge_2)==6) ):
            opgood="Awesome work! You have successfully completed atleast one track and are eligible for prizes. ^.^"
        #date
        f_date = dtt.today()
        l_date = dtt(2020, 11, 5)
        deltadate = l_date - f_date
        madate = str(deltadate)
        madate = madate[:2]
        renem1 =""
        renem2 = ""
        #print("sasds")
        #print (skbadgename_1)
        #print(skbadgename_2)

        if(str(skbadgename_1)=="None"):
            renem1="None"
        else:
            renem1 = skbadgename_1
        
        
        if(str(skbadgename_2)=="None"):
            renem2="None"
        else:
            renem2 = skbadgename_2
        #print("sdsds")
        
        

        #time
        from datetime import datetime
        now = datetime.now()

        current_time = now.strftime("%H:%M:%S")

        #print("Current Time =", current_time)
        h = int(now.strftime("%H"))
        m = int(now.strftime("%M")) 

            # Formula for total remaining minutes 
            # = 1440 - 60h - m 
        totalMin = 1440 - 60 * h - m 

            # Remaining hours 
        hoursRemaining = totalMin // 60

            # Remaining minutes 
        minRemaining = totalMin % 60
        minRemaining = minRemaining - 1

        print(hoursRemaining,"::",minRemaining) 

        # Driver code 

        # Current time 


        # Get the remaining time 
        #remainingTime(h, m) 

        


        mate = requests.post(
            "https://api.mailgun.net/v3/niran.dev/messages",
            auth=("api", "API KEY HERE"),
            
            data={"from": "[4 DAYS LEFT]-30DoGC <dscreva@niran.dev>",
                  "to": [stu_mail],
                  "subject": "Your Google Cloud Challenge progress report is here!".format(stu_name),
                  "template": "dsc",
                  "v:stu_name": stu_name,
                  "v:badstatus": opbad,
                  "v:goodstatus": opgood,
                  "v:dateleft": madate,
                  "v:hourleft": hoursRemaining,
                  "v:minleft": minRemaining,
                  "v:skbadge_1": skbadge_1,
                  "v:skbadgename_1": renem1,
                  "v:skbadge_2": skbadge_2,
                  "v:skbadgename_2": renem2 })
        #print(stu_name)
        print(stu_mail)
        print(mate)
        print("**********************")
        matte = str(mate)
        sheet1.write(cell, 1, stu_name) 
        sheet1.write(cell, 2, stu_mail) 
        sheet1.write(cell, 3, (skbadge_1+skbadge_2-1)) 
        sheet1.write(cell, 4, current_time) 
        sheet1.write(cell, 5, matte) 
        wb.save('logs3.xls')

        

        """if matte == "<Response [400]>" :
            print("Timeout, gotta wait for an hour")
            sheet1.write(cell, 6, "resend") 
            maddy = cell
            maddy = maddy = 1
            cell = maddy
            time.sleep(3600)
            
            print("imma awake ~")
        """    

        """
        print("****************************************")
        print("Sl. No.: {}".format(cell))
        print("Name: {}".format(stu_name))
        print("Mail: {}".format(stu_mail))
        print("SK Badge 1: {}".format(skbadge_1))
        print("SK Badge Name 1: {}".format(skbadgename_1))
        print("SK Badge 2: {}".format(skbadge_2))
        print("SK Badge Name 2: {}".format(skbadgename_2))
        print("****************************************")
        data = {
            "phone": stu,
            "body": ""
        }
        res = requests.post(url, json=data)
        print(res.text)
        print()
        print("Current count:", i)
        i += 1
        """

        # time.sleep(3)


targets = readContacts("mmm.xlsx")
