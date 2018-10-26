import xlrd
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

emailfrom ="DL-windstream_metasolv@Prodapt.com"
#emailto = 'balaji.ma@prodapt.com'
#emailto ='annamalai.d@prodapt.com'
#emailfrom = 'balaji.ma@prodapt.com'
emailto = 'DL-windstream_metasolv@prodapt.com'
#fileToSend = filename
date='10/12/2018'
msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = emailto
msg["Subject"] = "CRQ Scheduled on " + date 
msg.preamble = "CRQ Report"

book=xlrd.open_workbook('D:/Python/Automation tool/CRQ Status/Working code/CRQ Status.xls')
sheet=book.sheet_by_index(0)
print (sheet.nrows)
total_rows=sheet.nrows
total_cols=sheet.ncols
rownum1=[]
"""for i in range(1,total_rows):
    if sheet.cell_value(i,0) != '':
        print (i)
        rownum1.append(i)"""
for i in range(1,total_rows):
     if sheet.cell_value(i,0) != sheet.cell_value(i-1,0):
         print (i)
         rownum1.append(i)
print (rownum1)

html="""
<html>
<head>
<style>
  * { 
    margin: 0; 
    padding: 0; 
  }
  body { 
   font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif;
font-size: 16px;
font-style: normal;
font-variant: normal;
font-weight: 400;
line-height: 20px; 
  }
  .top {
  border-top: 2px solid black;
  border-color: black;
}
.bottom {
  border-bottom: 2px solid black;
  border-color: black;
}
.left {
  border-left: 2px solid black;
  border-color: black;
}
.right {
  border-right: 2px solid black;
  border-color: black;
}
  table { 
     border: 2px solid black;
    width: Auto; 
    border-collapse: collapse; 
  }
  tr:nth-of-type(odd) { 
    background: #eee; 
  }
  tr:nth-of-type(even) { 
    background: #fff; 
}
  th { 
background: #333; 
color: white; 
font-weight: bold; 
text-align: center; 
padding: 6px; 
border: 1px solid #9B9B9B; 
  }
  td { 
    padding: 6px; 
    border: 1px solid #9B9B9B; 
    text-align: center; 
  }
p {
     font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif;
font-size: 16px;
font-style: normal;
font-variant: normal;
font-weight: 400;
line-height: 20px;
}
</style>
</head>
<body>
<p>
Hi All ,
<br />
<br />
Please find the CRQ Status Report for date """+ date + """
<br />
<br />
</p>
<table border='1'>
<tr>
<th>CRQ</th>
<th>Owner</th>
<th>Status</th>
<th>Approver Status</th>
<th>Approver Group Name</th>
<th>Approver Names</th></tr>"""

#rownum1=[1,4,8,14,15,18]
print(len(rownum1))
for i in range(0,len(rownum1),1) :
    try:
        diff=rownum1[i+1]-rownum1[i]
        start=rownum1[i]
        end=rownum1[i+1]
    except :
        diff=(total_rows)-rownum1[i]
        start=rownum1[i]
        end=total_rows
        
    html=html+'<tr >' +"<td rowspan="+str(diff)+">"+sheet.cell_value(rownum1[i],0)+'</td>' +"<td rowspan="+str(diff)+">"+sheet.cell_value(rownum1[i],1)+'</td>' +"<td rowspan="+str(diff)+">"+sheet.cell_value(rownum1[i],2)+'</td>'        
    if diff > 1:
            for r in range(start,end,1):
                for col in range(3,6,1):
                    if sheet.cell_value(r,col)== 'Pending':
                        html=html+'<td style = "color:Orange">'+str(sheet.cell_value(r,col))+'</td>'
                    elif sheet.cell_value(r,col)== 'Approved':
                        html=html+'<td style = "color:Green">'+str(sheet.cell_value(r,col))+'</td>'
                    elif sheet.cell_value(r,col)== 'Rejected':
                        html=html+'<td style = "color:Red">'+str(sheet.cell_value(r,col))+'</td>'
                    else:
                        html=html+'<td HEIGHT=20">'+str(sheet.cell_value(r,col))+'</td>'

                    tr_var=1
                html = html + "</tr>"
    else:
        for col in range(3,6,1):
            html=html+'<td HEIGHT=20></td>'
        html = html + "</tr>"

html = html + "</table></body></html>"

#print (html)
                
            
part2 = MIMEText(html, 'html')
msg.attach(part2)
server = smtplib.SMTP("outlook.prodapt.com:587")
server.connect('outlook.prodapt.com')
server.sendmail(emailfrom, emailto, msg.as_string())
server.quit()
"""server = smtplib.SMTP("mailhost.windstream.com")
server.connect('mailhost.windstream.com')
server.sendmail(emailfrom, emailto, msg.as_string())
server.quit()"""
