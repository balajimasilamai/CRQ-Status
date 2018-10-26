import sqlite3
conn=sqlite3.connect('crq_status_database.db')

cur=conn.cursor()

cur.execute("""Create table if not exists crq_status (crq_number varchar2(50),
                                        Owner  varchar2(50),
                                        crq_status varchar2(50),
                                        task_status  varchar2(30),
                                        Approver_Group_Name varchar2(50),
                                        Approvers_Name varchar2(200),
                                        Approver_Sign varchar2(50),
                                        Approver_Alternate varchar2(50),
                                        Approval_date date)
                                       """)
#cur.execute("""insert into crq_status values ('CRQ000000118910','Annamalai Dakshinamurthy','Scheduled For Approval','Pending', 'Change Management','Michael A Patty;Jyotsna Koganti;William Masi;Amrita Newaskar',null,null,null)""")	
#cur.execute("""insert into crq_status values ('CRQ000000118910','Annamalai Dakshinamurthy','Scheduled For Approval','Approved','IT - Command Center','Marc Holtz;Steven Werner;Ranjitha Umapathy;Nandhini Arumugam;Anandavel Ramu;Abilash Sugavanam;Navinraj R;Janarthanan Vt;Pavithra Srinivasan;Suresh Vasudevan;Naresh Arumugam;Renukadevi V;Praveena Puttum;Sudheer K Kosgi;Mohamed Sarjoon;Adhavan Arumughaperumal;Vikas Govindraj;Mithra S Devi;Rajesh Loganathan;John F Morrissey;Hamkumar Sampath','Marc Holtz','Felix Liverman','10/9/2018 10:23:50 AM')""")	
#cur.execute("""insert into crq_status values ('CRQ000000118910','Annamalai Dakshinamurthy','Scheduled For Approval','Approved','IT-OSS - M6 (NextGen)','Michael J Sheriff;Chawn S Thompson;Bryan K Lewis;Deborah Philps;Dirk L Fox;Suriya S Kanthan',	'Michael J Sheriff',null,'10/9/2018 2:54:06 PM')""")	
#cur.execute("""insert into crq_status values ('CRQ000000118910','Annamalai Dakshinamurthy','Scheduled For Approval','Approved','DBA: Core','Timothy Brotzman;Jonathan D Mazak;Louis G Slavik;Harikrishna Rao;Joseph B Robinson Iii;Donna Menyes;Anand Buldeo',	'Anand Buldeo',null,'10/9/2018 2:19:12 PM')""")	
#cur.execute("""insert into crq_status values ('CRQ000000118809','Annamalai Dakshinamurthy','Scheduled For Approval','Pending', 'Change Management','Michael A Patty;Jyotsna Koganti;William Masi;Amrita Newaskar',null,null,null)""")				
#cur.execute("""insert into crq_status values ('CRQ000000118809','Annamalai Dakshinamurthy','Scheduled For Approval','Approved','IT - Command Center','Marc Holtz;Steven Werner;Ranjitha Umapathy;Nandhini Arumugam;Anandavel Ramu;Abilash Sugavanam;Navinraj R;Janarthanan Vt;Pavithra Srinivasan;Suresh Vasudevan;Naresh Arumugam;Renukadevi V;Praveena Puttum;Sudheer K Kosgi;Mohamed Sarjoon;Adhavan Arumughaperumal;Vikas Govindraj;Mithra S Devi;Rajesh Loganathan;John F Morrissey;Hamkumar Sampath'	,'Marc Holtz'	,'Felix Liverman'	,'10/9/2018 10:23:01 AM')""")	
#cur.execute("""insert into crq_status values ('CRQ000000118809','Annamalai Dakshinamurthy','Scheduled For Approval','Approved','IT-OSS - M6 (NextGen)','Michael J Sheriff;Chawn S Thompson;Bryan K Lewis;Deborah Philps;Dirk L Fox;Suriya S Kanthan','Michael J Sheriff'		,null,'10/9/2018 2:53:48 PM')""")	
#cur.execute("""insert into crq_status values ('CRQ000000118809','Annamalai Dakshinamurthy','Scheduled For Approval','Approved','DBA: Core','Timothy Brotzman;Jonathan D Mazak;Louis G Slavik;Harikrishna Rao;Joseph B Robinson Iii;Donna Menyes;Anand Buldeo',	'Anand Buldeo',null,		'10/9/2018 2:21:44 PM')""")	

#cur.execute('delete from crq_status ')
conn.commit()

op=cur.execute("""select crq_number,Approver_Group_Name,count(*)
                  from crq_status where crq_number='CRQ000000118910'
                  group by crq_number,Approver_Group_Name
               """)
#op=cur.execute("select *   from crq_status where crq_status='CRQ000000118910' ")
for i in op.fetchall():
   if i[2]>1:
       op1=cur.execute("select crq_number,Approver_Group_Name,max(Approval_date) from crq_status where crq_number=? and Approver_Group_Name=?",(i[0],i[1],))
       for date in op1.fetchall():
           print (date)
       
