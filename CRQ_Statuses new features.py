# importing the modules
import time
import win32com.client
import autoit
from selenium import webdriver
import csv
import pandas as pd
import re
from PIL import Image, ImageTk
#from HoverInfo import HoverInfo
from datetime import datetime
import sys
from tkinter import * 
from tkinter import messagebox
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
#from selenium.webdriver.chrome.options import Options
import xlwt
from threading import Thread
from tkinter import ttk
import pythoncom
#====================================
def on_entry_click(event):
    if Date_entry.get() == 'M/D/YYYY':
       Date_entry.delete(0, "end") # delete all the text in the entry
       Date_entry.insert(0, '')
       Date_entry.configure(fg='black')
#=================Funtion to send a mail
def send_mail(filename,to_addr):
    import smtplib
    import mimetypes
    import socks
    from email.mime.multipart import MIMEMultipart
    from email import encoders
    from email.message import Message
    from email.mime.audio import MIMEAudio
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email.mime.text import MIMEText
    socks.setdefaultproxy(socks.HTTP, 'proxy.windstream.com', 8080)
    socks.wrapmodule(smtplib)

    emailto = to_addr#"Balaji.Masilamani@windstream.com"
    emailfrom = 'CRQStatus@windstream.com'
    fileToSend = filename

    msg = MIMEMultipart()
    msg["From"] = emailfrom
    msg["To"] = emailto
    msg["Subject"] = "CRQ Status Report for date "+str(Date_entry.get())
    msg.preamble = "CRQ Status Report"

    ctype, encoding = mimetypes.guess_type(fileToSend)
    if ctype is None or encoding is not None:
     ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
     fp = open(fileToSend)
    # Note: we should handle calculating the charset
     attachment = MIMEText(fp.read(), _subtype=subtype)
     fp.close()
    elif maintype == "image":
     fp = open(fileToSend, "rb")
     attachment = MIMEImage(fp.read(), _subtype=subtype)
     fp.close()
    elif maintype == "audio":
     fp = open(fileToSend, "rb")
     attachment = MIMEAudio(fp.read(), _subtype=subtype)
     fp.close()
    else:
     #with open(fileToSend,  'r', encoding='latin-1') as fp:
     fp = open(fileToSend, "rb")
     attachment = MIMEBase(maintype, subtype)
     attachment.set_payload(fp.read())
     fp.close()
     encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
    msg.attach(attachment)

    server = smtplib.SMTP("mailhost.windstream.com")
    #server.starttls()
    server.connect("mailhost.windstream.com")
    #server.login(username,password)
    try:
        server.sendmail(emailfrom, emailto,  msg.as_string())
    except Exception as e:
        messagebox.showinfo('Error', str(e) + '\n' + 'Check To address')
    server.quit()

#============== Styles for Excel sheet =====================================
styleOK = xlwt.easyxf('pattern: fore_colour green;'
                          'font: colour black, bold True;'
                          'align: wrap yes,vert centre, horiz center;'
                 'borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
border=xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
alignments = xlwt.easyxf("align: wrap true, horiz center, vert center;"
                       'borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;')
#style_string = "font: bold on; borders: bottom dashed"
#style = xlwt.easyxf(style_string)
style = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;'
                              'font: colour white, bold True;'
                    'borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thick, bottom thin;')

#========== OPen the sheet to get the CRQ number=====================================================
listofcrq=[]
count=0
crq_found=0
listtime=[]

#===== threading function =============
def start_thread():
        t = Thread(target=open_browser)
        t.start()
# ==== Openening a browser ==========
def open_browser():

 if Date_entry.get() != 'M/D/YYYY' and to_address_entry.get() !='':
    date=Date_entry.get()
    itsm_query=''''Scheduled Start Date+'  >=  "'''+str(date)+'''"   AND    ( 'Service+' = "M6 (EarthLink)"  OR 'Service+' = "M6 (NextGen)"  OR 'Service+' = "M6 (PAETEC)"  OR 'Service+' = "M6 (ASAP/TSG)" )  AND  ( 'Status*' = "Draft"   OR 'Status*' = "Scheduled For Review"  OR 'Status*' = "Scheduled For Approval"  OR 'Status*' = "Planning In Progress"  OR 'Status*' = "Scheduled"  OR 'Status*' = "Implementation In Progress"  OR 'Status*' = "Rejected"  OR 'Status*' = "Cancelled")'''
    progessbar.grid(row=4,column=1,sticky=E,padx=10)
    progessbar.start()
    global count
    global crq_found
    print (Date_entry.get())
    driver=webdriver.Chrome()#chrome_options=options
    driver.maximize_window()
    try:
        driver.get('https://itsm.windstream.com/')
        time.sleep(30)
        wait(driver,60)
        pythoncom.CoInitialize()
        aw=True
        while aw:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.Sendkeys(username_entry.get())#n9941391#n9930786
            shell.Sendkeys('{TAB}')
            shell.Sendkeys(password_entry.get())#Jan2018$#MssCSO001@
            shell.Sendkeys('{ENTER}')
            aw=False
        
    except Exception as e:
        print (e)
        progessbar.grid_forget()
        messagebox.showinfo('Warning','Connect VPN/check your username and password')
        
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,30).until(alert_wait)
        alert=driver.switch_to_alert()
        alert.accept()
    except:
        pass
    start_time=time.strftime('%H:%M:%S')
    listtime.append(start_time)
    logo_present = EC.presence_of_element_located((By.ID, 'reg_img_304316340'))
    wait(driver ,50).until(logo_present)
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,10).until(alert_wait)
        alert=driver.switch_to_alert()
        alert.accept()
    except:
        pass
    level_element=driver.find_element_by_id('reg_img_304316340')
    level_element.click()
    wait(driver,20)
    child_element= wait(driver,15).until(EC.element_to_be_clickable((By.XPATH ,"//span[text()='Change Management']")))
    child_element.click()
    wait(driver,20)
    #Changed the code to click 'Search Change' instead of clicking 'Change management console'
    child_element_1=wait(driver,15).until(EC.element_to_be_clickable((By.XPATH ,"//span[text()='Search Change']")))
    child_element_1.click()
    wait(driver,25)
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,10).until(alert_wait)
        alert=driver.switch_to_alert()
        alert.accept()
    except:
        pass
    #Need to add the code for clicking advance search button using xpath='//*[@id="TBadvancedsearch"]'
    adv_button= wait(driver,100).until(EC.element_to_be_clickable((By.XPATH,"//fieldset/div/div/div/div[1]/table/tbody/tr/td[3]/a[3][@arwindowid='3']/div")))
    #html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[3]/a[3]/div
    if adv_button.is_displayed():    
    #adv_button= wait(driver,50).until(EC.element_to_be_selected((By.XPATH,'//*[@id="TBsavedsearches"]')))
     ActionChains(driver).double_click(adv_button).perform()
    else:
        print('Not displayed')
    #Need to add the code for sending ITSM query to text area
    wait(driver,50)
    text_area=driver.find_element_by_xpath('//fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[5]/table[2]/tbody/tr/td[1]/textarea[@id="arid1005"]')
    text_area.click()
    text_area.send_keys(itsm_query)
    #Need to add the code to click the search button
    wait(driver,5)
    search=driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[16]/div/div/div[4]/fieldset/div/div/div/div/div[2]/fieldset/div/a[2]/div/div')
    search.click()
    #Need to add the code to get the list of 
    
    '''drop_down=wait(driver,10).until(EC.element_to_be_clickable((By.XPATH ,'//*[@id="WIN_3_303174300"]/div/a/img')))
    drop_down.click()
    wait(driver,10)
    choose_drop_down=driver.find_element_by_xpath('html/body/div[3]/div[2]/table/tbody/tr[5]/td[1]').click()'''
    wait(driver,10)
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,8).until(alert_wait)
        alert=driver.switch_to_alert()
        alert.accept()
    except:
        pass
    try:
        wait(driver,10)
        popup=driver.find_element_by_xpath('//*[@id="PopupMsgFooter"]/a')
        popup.click()
    except:
        pass
    #========Geting the CRQ ids ==========================================
    #Need to modify the code to get the list of crqs
    wait(driver,30)
    CRQ_xpath=driver.find_elements_by_xpath(".//*[@id='T1020']/tbody/tr")
    print (CRQ_xpath)
    print (len(CRQ_xpath))
    progessbar.configure(maximum=len(CRQ_xpath))
    number=0
    output=''
    outer=0
    r=1
    c=0
    column=0
    top_row = 1
    bottom_row=0
    left_column = 0
    right_column = 0
    #list_status=['/html/body/div[4]/div[2]/table/tbody/tr[1]/td[2]',
            # '/html/body/div[4]/div[2]/table/tbody/tr[2]/td[2]',
            # '/html/body/div[4]/div[2]/table/tbody/tr[3]/td[2]']

    columns=['CRQ Number','Owner','CRQ Status','Task Status','Approver Group Name','Approvers Name','Approver Sign','Approver Alternate','Approval date']
    approved_status=False
    # OPeneing a new Excel sheet to load the data
    filename='CRQ Status.xls'
    wb = xlwt.Workbook()
    ws = wb.add_sheet('CRQ_STATUS',cell_overwrite_ok=True)

    # to add the headers
    for i in columns:
        ws.write(0,column,i,style)
        column+=1
    ws.col(0).width =256 * 20
    ws.col(1).width =256 * 35
    ws.col(2).width =256 * 30
    ws.col(3).width =256 * 15
    ws.col(4).width =256 * 20
    ws.col(5).width =256 * 70
    ws.col(6).width =256 * 70
    ws.col(7).width =256 * 70
    ws.col(8).width =256 * 70
    start=2
    for i in range(0,len(CRQ_xpath),1):
      wait(driver,40)
      print ('Loop: '+str(i)+'   '+'element '+str(start))
      try:
                  crq_start=time.strftime('%H:%M:%S')
                  listtime.append(crq_start)
                  crq_found=crq_found+1
                  print (CRQ_xpath[i].text)
                  wait(driver,30)
                  #element=CRQ_xpath[i]
                  element=driver.find_element_by_xpath("//*[@id='T1020']/tbody/tr["+str(start)+"]")
                  #element=ActionChains(driver).double_click(element).perform() # NO need of double click chage it to Single click
                  wait(driver,10)
                  element.click()
                  wait(driver,20)
                  try:
                      alert_wait=EC.alert_is_present()
                      wait(driver ,30).until(alert_wait)
                      alert=driver.switch_to_alert()
                      alert.accept()
                  except:
                      pass
                  wait(driver,20)
                  time.sleep(3)
                  change_id=driver.find_element_by_id("arid_WIN_3_1000000182")
                  crq_value=change_id.get_attribute('value')
                  #ws.write(r, c,change_id.get_attribute('value'),alignments)
                  wait(driver,5)
                  owner=driver.find_element_by_id("arid_WIN_3_1000003230")
                  owner_value=owner.get_attribute('value')
                  #ws.write(r, c+1,owner.get_attribute('value'),alignments)
                  wait(driver,2)
                  crq_status=driver.find_element_by_id('arid_WIN_3_303502600')
                  status_crq=crq_status.get_attribute('value')
                  #ws.write(r, c+2,crq_status.get_attribute('value'),alignments)
                  wait(driver,5)
                  for status in range(1,4,1):
                        drop_element=wait(driver,10).until(EC.element_to_be_clickable((By.XPATH ,'/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[16]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[6]/fieldset/div/div[3]/div/a/img')))
                        drop_element.click()
                        wait(driver,45)
                        #time.sleep(5)
                        choose=driver.find_element_by_xpath('/html/body/div[4]/div[2]/table/tbody/tr['+str(status)+']/td[1]').click()
                        wait(driver,45)
                        #time.sleep(6)
                        Tech_review=driver.find_elements_by_xpath("./html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[16]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[7]/fieldset/div/div/div[2]/div/div[2]/table/tbody/tr")
                        wait(driver,5)
                        if len(Tech_review)> 0:
                            count=1
                            for i in range (2,len(Tech_review)+1,1):
                                ws.write(r, c,str(crq_value),alignments)
                                ws.write(r, c+1,str(owner_value),alignments)
                                ws.write(r, c+2,str(status_crq),alignments)
                                for i1 in range(1,7,1):
                                    wait(driver,5)
                                    element_text=driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[16]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[7]/fieldset/div/div/div[2]/div/div[2]/table/tbody/tr["+str(i)+']/td['+str(i1)+']/nobr/span')
                                    ws.write(r, i1+2,str(element_text.text),border)
                                    i1=i1+1
                                
                                bottom_row=bottom_row+1
                                r=r+1
                  """if len(Tech_review)==0:
                                ws.write(r, c,str(crq_value),alignments)
                                ws.write(r, c+1,str(owner_value),alignments)
                                ws.write(r, c+2,str(status_crq),alignments)
                                r=r+1"""
        
                  if count > 0:
                      count=0
                      #ws.write_merge(top_row,bottom_row, left_column, right_column, change_id.get_attribute('value'),alignments)
                      #ws.write_merge(top_row,bottom_row, left_column+1, right_column+1, owner.get_attribute('value'),alignments)
                      #ws.write_merge(top_row,bottom_row, left_column+2, right_column+2, crq_status.get_attribute('value'),alignments)
                      top_row=r
                      wb.save(filename)
                  else:
                      ws.write(r, c,str(crq_value),alignments)
                      ws.write(r, c+1,str(owner_value),alignments)
                      ws.write(r, c+2,str(status_crq),alignments)
                      ws.write(r, c+3,None,border)
                      ws.write(r, c+4,None,border)
                      ws.write(r, c+5,None,border)
                      r=r+1
                      top_row=r
                      bottom_row=r-1
                      wb.save(filename)
                      
                  wait(driver,5)
                  #back=driver.find_element_by_id("reg_img_304248620")
                  #back.click()
                  #wait(driver,8)
                  crq_end=time.strftime('%H:%M:%S')
                  listtime.append(str(crq_end))
                  start=start+1
      except Exception as e:
            print (e)
            pass
            #progessbar.grid_forget()
            #messagebox.showinfo('Error','Not able to find the element due to browser loading issue or timeout'+'\n'+'Please initiate the process again')
      var.set(i+1)
      end_process=time.strftime('%H:%M:%S')
      listtime.append(str(end_process))
    #wb.save(filename)
    if crq_found>0:
        send_mail(filename,to_address_entry.get())
    else:
        progessbar.grid_forget()
        messagebox.showinfo('Warning','No CRQ found for your search')
        
    time.sleep(5)
    logout=driver.find_element_by_xpath('//*[@id="WIN_0_300000044"]/div/div')
    logout.click()
    driver.quit
    progessbar.grid_forget()
 elif Date_entry.get() != 'M/D/YYYY' and to_address_entry.get() =='':
     messagebox.showinfo('Warning','Enter the To email address to proceed further')
 elif Date_entry.get() == 'M/D/YYYY' and to_address_entry.get() !='':
     messagebox.showinfo('Warning','Enter the DATE to proceed further')
 else:
     messagebox.showinfo('Warning','Enter the DATE & to email Address to proceed further')
 time_calculation()
     
def time_calculation():
    t=2
    i1=1
    print ('Found all crq and starting Process ',listtime[0])
    for i in range(0,int((len(listtime)-2)/2)):        
        FMT = '%H:%M:%S'
        time_diff = datetime.strptime(listtime[t], FMT) - datetime.strptime(listtime[i1], FMT)
        print ('CRQ'+str(i+1)+': ' , time_diff) 
        t=t+2
        i1=i1+2
    print ('Process End time ',listtime[len(listtime)-1])
import PIL
#creating the loging info page
root=Tk()
s = ttk.Style()
s.theme_use('classic')
s.configure("blue.Horizontal.TProgressbar", foreground='blue', background='blue')
root.geometry("600x300")
root.title('CRQ Status')
root.resizable(False, False)
image = PIL.Image.open('images.jpg')
photo_image = ImageTk.PhotoImage(image)
l = Label(root, image = photo_image)   
#l.place(x=18,y=-1,relwidth=1, relheight=1)
root.configure(background='sky blue')
username=Label(root,text='User Name:',
               #bg='sky blue',
               font=('Arial',10,'bold'))
username.grid(row=0,column=0,sticky=W,padx=100,pady=25)
username_entry=Entry(root,bd=5,width=30)
username_entry.insert(END,'n9941391')
username_entry.grid(row=0,column=1,sticky=W)
password=Label(root,text='Password:',
               #bg='sky blue',
               font=('Arial',10,'bold'))
password.grid(row=1,column=0,sticky=E,pady=10,padx=100)
password_entry=Entry(root,bd=5,show='*',width=30)
password_entry.insert(END,'Sep2018$')
password_entry.grid(row=1,column=1,sticky=W)
Date=Label(root,text='Date:',
           #bg='sky blue',
           font=('Arial',10,'bold'))
Date.grid(row=2,column=0,sticky=E,pady=10,padx=100)
Date_entry=Entry(root,bd=5,width=30,fg='grey')
Date_entry.grid(row=2,column=1,sticky=W)
Date_entry.insert(END,'M/D/YYYY')
Date_entry.bind('<Key>', on_entry_click)
#hover=HoverInfo(Date_entry,'M/D/YYYY')
#Dateexample=Label(root,text='Eg:M/D/YYYY',bg='sky blue')
#Dateexample.grid(row=2,column=2)
to_address=Label(root,text='Email To:',
                 #bg='sky blue',
                 font=('Arial',10,'bold'))
to_address.grid(row=3,column=0,sticky=E,pady=10,padx=100)
to_address_entry=Entry(root,bd=5,width=30)
to_address_entry.grid(row=3,column=1,sticky=W)
to_address_entry.insert(END,'n9996094@windstream.com')
var = IntVar()
progessbar = ttk.Progressbar (root, variable=var, orient='horizontal', length=200,style="red.Horizontal.TProgressbar")
progessbar.grid(row=4,column=1,sticky=E,padx=10)
var.set(0)
progessbar.grid_forget()
button=Button(root,text='Get CRQ status',bg='sky blue',command=start_thread)
button.grid(row=5,column=1,pady=10)
button.bind('<Return>',start_thread)
root.mainloop()




