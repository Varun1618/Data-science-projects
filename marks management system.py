#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xlrd
import numpy as np
import matplotlib.pyplot as plt
def mail(f,t,s):
    import smtplib
    import os
    from email import encoders
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email import encoders
    sender='bachuvarun123@gmail.com'
    receiver=t
    sub=s
    message=MIMEMultipart()
    message['from']=sender
    message['to']=receiver
    message['subject']=sub
    msg_content='This is the report'
    message.attach(MIMEText(msg_content,'plain'))
    attachment=open(f,'rb')
    msg=MIMEBase('application','octet-stream')
    msg.set_payload(attachment.read())
    encoders.encode_base64(msg)
    msg.add_header('Content-Disposition','attachment',filename=os.path.basename(f))
    message.attach(msg)
    text=message.as_string()
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(sender,'Varun2202@')
    server.sendmail(sender,receiver,text)
    server.close()
print("Welcome to Marks Management System")
path="C:/Users/Asus/Desktop/studentmarks.xlsx"
workbook=xlrd.open_workbook(path)
sheet1=workbook.sheet_by_index(0)
sheet2=workbook.sheet_by_index(1)
print("Enter 1 for student wise marks...")
print("Enter 2 for subject wise marks...")
print("Enter 3 for Subject wise Comparision...")
print("Enter 4 to send reports to all parents...")
print("Enter 5 to send reports to the parents of weak students...")
print("Enter 6 to send subject wise reports to the respective subject teachers...")
print("Enter 7 for exit...")
while(True):
    a=int(input("Enter the digit to select from the above menu: "))
    if(a==1):
        print("Student wise results...")
        lst=[sheet1.cell_value(i,0) for i in range(1,sheet1.nrows)]
        studentname=input("Enter the student name: ")
        if studentname in lst:
            n=lst.index(studentname)
            n+=1
            x1=[2,4,6,8,10,12]
            y1=[sheet1.cell_value(n,i) for i in range(2,7)]
            y1.append(sheet1.cell_value(n,8))
            y2=[35,35,35,35,35,35]
            y3=[60,60,60,60,60,60]
            labels=['Maths','Physics','Chemistry','Biology','English',' Percentage']
            plt.bar(x1,y1,tick_label=labels,color='blue',width=0.8)
            plt.plot(x1,y2,linestyle='dashed',label='Fail',color='red')
            plt.plot(x1,y3,linestyle='dashed',color='green',label='weak')
            plt.ylim(0,100)
            plt.xlabel('Subjects')
            plt.ylabel('Marks')
            plt.title('Results of %s'%studentname)
            for j in range(6):
                plt.text(x1[j],y1[j],'%.1f'%y1[j])
            plt.legend()
            plt.show()
            plt.close()
        else:
            print("Check the Student name! Enter again...")
    elif(a==2):
        print("Subject Wise Results...")
        lst1=[sheet1.cell_value(0,i) for i in range(2,7)]
        subject=input("Enter the subject name: ")
        if subject in lst1:
            n1=lst1.index(subject)
            n1+=2
            x=[]
            for i in range((sheet1.nrows)-1):
                i+=1
                x.append(i+1)
            y=[sheet1.cell_value(i,n1) for i in range(1,sheet1.nrows)]
            y2=[35,35,35,35,35,35]
            y3=[60,60,60,60,60,60]
            labels=[sheet1.cell_value(k,0) for k in range(1,7)]
            plt.bar(x,y,tick_label=labels,color='blue',width=0.4)
            plt.plot(x,y2,linestyle='dashed',color='red',label='Fail')
            plt.plot(x,y3,linestyle='dashed',color='green',label='weak')
            plt.ylim(0,100)
            plt.xlabel('Students')
            plt.ylabel('Marks')
            plt.title('Marks of %s' %subject)
            for i in range(5):
                plt.text(x[i],y[i],'%.1f'%y[i])
            plt.legend()
            plt.show()
            plt.close()
        else:
            print("Check the subject name! Try again....")
    elif(a==3):
        print("Subject wise comparision...")
        x=[]
        for i in range(2,7):
            x.append(i)
        y1=[]
        for j in range(2,7):
            y=[sheet1.cell_value(i+1,j) for i in range((sheet1.nrows)-1)]
            d=np.array(y)
            b=np.mean(d)
            y1.append(b)
        labels=[sheet1.cell_value(0,i) for i in range(2,7)]
        plt.bar(x,y1,tick_label=labels,color='blue',width=0.4)
        plt.ylim=(0,100)
        plt.xlabel('Subjects')
        plt.ylabel('Marks')
        plt.title('Subject wise marks comparision')
        for i in range(5):
            plt.text(x[i],y1[i],'%.1f'%y1[i])
        plt.show()
        plt.close()
    elif(a==4):
        print("Send Reports to all parents....")
        lst3=[sheet1.cell_value(i,0) for i in range(1,sheet1.nrows)]
        for student in lst3:
            m=lst3.index(student)
            x=[]
            for i in range(2,sheet1.ncols):
                x.append(i)
            y=[sheet1.cell_value(m+1,i) for i in range(2,sheet1.ncols)]
            labels=[sheet1.cell_value(0,j) for j in range(2,sheet1.ncols)]
            plt.bar(x,y,tick_label=labels,color='blue',width=0.4)
            plt.ylim(0,400)
            plt.xlabel('Subjects')
            plt.ylabel('Marks')
            plt.title("Report of %s"%student)
            plt.savefig('report.png')
            for i in range(5):
                plt.text(x[i],y[i],'%.1f'%y[i])
            mail('report.png',sheet1.cell_value(m+1,1),'Report of your Son/Daughter')
            plt.close()
        print("All Emails sent...")
    elif(a==5):
        print("Send reports to the parents of weak students....")
        lst4=[sheet1.cell_value(i,0) for i in range(1,sheet1.nrows)]
        for student in lst4:
            s=lst4.index(student)
            s+=1
            if(sheet1.cell_value(s,8)<60):
                x1=[]
                for i in range(2,sheet1.ncols):
                    x1.append(i)
                y1=[sheet1.cell_value(s,j) for j in range(2,sheet1.ncols)]
                labels1=[sheet1.cell_value(0,k) for k in range(2,sheet1.ncols)]
                plt.bar(x1,y1,tick_label=labels1,color='blue',width=0.4)
                plt.ylim(0,400)
                plt.xlabel('Subjects')
                plt.ylabel('Marks')
                plt.title('Report of %s'%student)
                for i in range(7):
                    plt.text(x1[i],y1[i],'%.1f'%y1[i])
                plt.savefig('report1.png')
                mail('report1.png',sheet1.cell_value(s,1),'Report of your son/daughter')
                plt.close()
        print("All Emails sent....")
    elif(a==6):
        print("Send subject wise report to the respective subject teacher....")
        lst5=[sheet2.cell_value(j,0) for j in range(1,sheet2.nrows)]
        lst6=[sheet2.cell_value(k,1) for k in range(1,sheet2.nrows)]
        for sub in lst5:
            s1=lst5.index(sub)
            s1+=1
            x2=[]
            for i in range(1,sheet1.nrows):
                x2.append(i)
            y2=[sheet1.cell_value(j,s1+1) for j in range(1,sheet1.nrows)]
            labels2=[sheet1.cell_value(i,0) for i in range(1, sheet1.nrows)]
            plt.bar(x2,y2,tick_label=labels2,color='blue',width=0.4)
            plt.ylim(0,100)
            plt.xlabel('Students')
            plt.ylabel('Marks')
            plt.title('Report of %s'%sub)
            for i in range(6):
                plt.text(x2[i],y2[i],'%.1f'%y2[i])
            plt.savefig('subject report.png')
            mail('subject report.png',sheet2.cell_value(s1,2),'Subject analysis report')
        print("All Emails sent....")
    elif(a==7):
        break
        
            


# In[ ]:





# In[ ]:




