import xlrd
import xlwt
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import dateutil.parser
import datetime
import win32com.client
rownum=[]
r_list=[]
delete_row=[]
delete_row1=[]
def read_execl():
     book=xlrd.open_workbook('D:/Python/Automation tool/CRQ Status/Working code/CRQ Status.xls', on_demand = True)
     sheet=book.sheet_by_index(0)
     total_rows=sheet.nrows
     total_cols=sheet.ncols
     row=0
     for i in range(1,total_rows):
          if sheet.cell_value(i,0) != sheet.cell_value(i-1,0):
              #print (i)
              rownum.append(i)
     #print (rownum)#1,5


     r=0
     for i in range(0,len(rownum),1):
              print (' start row number  '+str(i))
              if rownum[i] != rownum[-1] :
                  diff = rownum[i+1] -rownum[i]
                  start=rownum[i]
                  end=rownum[i+1]
                  #print ('dif.....'+str(diff))
                  #print ('start row..',start)
                  #print ('end row...',end)
              else:
                  diff = total_rows - rownum[i]
                  start=rownum[i]
                  end=total_rows        
                  #print ('except')
                  #print ('dif.....'+str(diff))
                  #print (start)
                  #print (end)
              if diff > 1:
                    for row in range(start,end,1):
                     r=row
                     for i in range(start,end):
            
                         try:
                               if row != i and sheet.cell_value(row,4) == sheet.cell_value(i,4):
                                   #print ('row.....'+str(row))
                                   #print (str(row) +' value '+sheet.cell_value(row,4) +' '+ str(i) +' value '+sheet.cell_value(i,4))
                                   delete_row.append((row,i))
                                   #print (sheet.cell_value(row,4))
                                   #print (sheet.cell_value(i,4))
                                   #print ('end row.....'+str(r))
                                   r=r+1
                               if r == end:
                                   r=end-1
                         except:
                                   pass
     #print ('delete_row',delete_row)




     delete_row2=[]
     #rint ('length of deleterow ',len(delete_row))
     row=[]
     for i in range(0,len(delete_row),1):
          #print (i)
          for j in range(1,len(delete_row)):
               if i not in row and delete_row[i][0]==delete_row[j][1] and delete_row[i][1]==delete_row[j][0]:
                    #print (delete_row1[i])
                    delete_row2.append(delete_row[i])
                    row.append(j)
     print (delete_row2)
     for i in delete_row2:
          delete_row.remove(i)
          #print ('delete_row',delete_row)
     
     for i in delete_row:
          try:
              #print (i[0])
              #print (i[1])
              r=int(i[0])
              r1=int(i[1])

              parsed1 = dateutil.parser.parse(sheet.cell_value(r,8))
              parsed2 = dateutil.parser.parse(sheet.cell_value(r1,8))
              #print (parsed1)
              #print (parsed2)
              if parsed1 >  parsed2: 
                r_list.append(r1)
                #print (r1)
                
              elif parsed1 == parsed2 :
                r_list.append(r1)
                #print (r1)
              elif sheet.cell_value(r,8) is None and sheet.cell_value(r1,8) is not None  :
                r_list.append(r1)
                #print (r1)
              elif sheet.cell_value(r,8) is not None and sheet.cell_value(r,8) is  None  :
                r_list.append(r)
                #print (r)
              else:                
                 r_list.append(r)
                 #print (r)
          except:
                    pass
     print ('r_list---',r_list)

     book.release_resources()
     #del sheets
     del book

read_execl() 
if len(r_list)> 0:
       print ('legnth of r_list',len(r_list))
       for i in range(0,1,1):              
          filename='D:/Python/Automation tool/CRQ Status/Working code/CRQ Status.xls'
          app = win32com.client.Dispatch("Excel.Application")
          book=app.Workbooks.Open(Filename=filename)
          active_sheet=book.ActiveSheet
          print (active_sheet.UsedRange.Rows.Count)     
          line=r_list[i]
          active_sheet.Rows(line+1).EntireRow.Delete()
          print ('Deleted one row')
          #book.SaveAs(filename)
          book.Close()
          app.Quit()
else:
          start=False
          print ('No duplicate rows found')
     
