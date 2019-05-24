import time
from selenium import webdriver
import re
from datetime import datetime, timedelta
import pandas as pd
from pandas import dateFrame
from collections import defaultdict
from flashtext import KeywordProcessor

#change name of file where you want to save your results, month of the year, username and password

url_log = "https://login.sling.is/"
url=""

sum_min=datetime.strptime('0000', '%H%M')
hour=0
workers_nr=1
writer = pd.ExcelWriter(XLSX_FILE)
month=MONTH_NAME

driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
driver.get(url_log)

username = driver.find_element_by_name("emailAddress")
password = driver.find_element_by_name("password")
username.send_keys(USERNAME)
password.send_keys(PASSWORD)

driver.find_element_by_class_name("progress-button-content-original").click()

time.sleep(15)
url=driver.current_url
print(url)

driver.get(url+"shifts?mode=month&date="+DATA+"&tab=alllocations") 
time.sleep(7) 
print(driver.current_url)
time.sleep(2) 
driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/button').click()
position_list=driver.find_element_by_xpath('//*[@id="content"]/div/div[2]/div/shifts-sidebar/div/div[2]/div/sidebar-filter[2]/div/div[2]').text.split("\n")


#############################################################################################################################################################
names_list=[]

if url != url_log:
    
    time.sleep(2)
    EN=driver.find_element_by_xpath('//*[@id="shifts-main-container"]/div[2]/div[2]/div/div/div/table/tbody').text
    HALO=EN
    EN=re.split('\d \w+',EN)[1]
    text=EN[1:].split('\n')
    
    nam_list = [x for x in text if not (x.isdigit())]
    for i in range(2, len(nam_list),3):
        names_list.append(nam_list[i])
        
    names_list=list((set(names_list)))
    
    for i in range(0,len(names_list)):
        text=''
        lista=[]
        listaa=[]
        absence=0
        sm=0
        for j in range(0,len(nam_list)):
            if names_list[i]==nam_list[j]:
                print(names_list[i])
                text=text+str(nam_list[j-2]) +'\n' +str(nam_list[j-1])+'\n'
                listaa=text.split('\n')
                absence = absence+listaa.count('All day')
                text=text.replace('•', '')
                text=text.replace('TEST', '')
                text=text.replace('All day', '')
                text=text.replace('Time off','')
                text=text.replace('P  ', 'P\n')
                text=text.replace('A  ', 'A\n')
                text=re.sub('(?<=\d)min',r'', text)
                text=re.sub('(?<=\d)h ?',r'', text)

#############################################################################################################################################################
        
        if text.isspace()==False:
            lista=text.split("\n")
            lista = list(filter(None, lista))
            hours=[]
            numbers=[]
            activity=[]
            date=[]
            
            for k in range(0,len(lista),3):
                     print(lista[k]+' '+lista[k+1]+' '+lista[k+2])
                     hours.append(lista[k])
                     numbers.append(lista[k+1])
                     activity.append(lista[k+2])
                
            for k in range(0,len(numbers)):   
                    if len(numbers[k])==3:
                        numbers[k]='0'+numbers[k]
                    elif len(numbers[k])==1:
                        numbers[k]='0'+numbers[k] +'00'
                    elif len(numbers[k])==2:
                        numbers[k]='00'+numbers[k]
                    date.append(datetime.strptime(numbers[k], '%H%M'))
        else:
            lista=[]
            hours=[]
            numbers=[]
            activity=[]
            date=[]
###############################################################################################################################################################        

        save_list=[]
        time_list=[]
        absence_list=[]
        lista_procent=[]
        lista_procent2=[]
        

        minute=datetime.strptime('0000', '%H%M')
        

        for p in range (0,len(position_list)):
            for q in range(0,len(activity)):
                if position_list[p]==activity[q]:
                   sum_min=sum_min + timedelta(minutes=int(datetime.strftime(date[q], '%M')))
                   minute=datetime.strftime(sum_min, '%M')
                   
                   hour=hour+int(datetime.strftime(date[q], '%H'))
            hour_zczyt = datetime.strftime(sum_min, '%H')
            hour=hour+int(hour_zczyt)
            
            if str(minute)!='1900-01-01 00:00:00':
                print(position_list[p]+" "+str(hour)+":"+str(minute))
                
                save_list.append(position_list[p])
                a=[]
                
                time_list.append(str(hour)+" h "+str(minute)+" min")
                
                if str(minute) == '15':
                    lista_procent.append(float(str(hour)+"."+'25'))
                elif str(minute) == '30':
                    lista_procent.append(float(str(hour)+"."+'5'))
                elif str(minute) == '45':
                    lista_procent.append(float(str(hour)+"."+'75'))
                else:
                    lista_procent.append(float(str(hour)))
                
                sm = sum(lista_procent[0:len(lista_procent)])
                
                
                
                absence_list=[absence]
                lista_month=[month]
                lista_imienia=[names_list[i]]
                

                for y in range (1, len(save_list)):
                        absence_list.append('-')
                        lista_month.append(month)
                        lista_imienia.append(names_list[i])
                
                               
                df = dateFrame({'Time': time_list,'Sling': save_list, 'Absence':absence_list, 'Month': lista_month, 'Person': lista_imienia})
                df=df[['Sling','Time','Absence','Month', 'Person']]
                df.to_excel(writer, sheet_name=names_list[i], index=False)
                

                workbook  = writer.book
                worksheet = writer.sheets[names_list[i]]
                format2 = workbook.add_format({'align': 'center'})
                worksheet.set_column('A:A', 50, None)
                worksheet.set_column('B:B', 13, format2)
                worksheet.set_column('C:C', 50, None)
                worksheet.set_column('D:D', 15, format2)
                worksheet.set_column('F:F', 20, None)
                worksheet.set_column('G:G', 20, None)


            sum_min=datetime.strptime('0000', '%H%M')
            minute=datetime.strptime('0000', '%H%M')
            hour_zczyt='00'
            hour=0
            
    
        print("-----------------------------------------------------------------------------------------")
        
        
        
print("DONE")
writer.save()   
#        