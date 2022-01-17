

import xlrd
import xlutils
path="C:/Users/Asus/Desktop/data10.xlsx"
w=xlrd.open_workbook(path)
sheet=w.sheet_by_index(0)
lst=[sheet.cell_value(k,0) for k in range(1,sheet.nrows)]





import xlwt
import xlrd
from xlutils.copy import copy 
path1="C:/Users/Asus/Desktop/data11.xls"
w1=xlrd.open_workbook(path1)
sheet2=w1.sheet_by_index(0)
print("Welcome to Techvanto Bank")
print("Enter 1 for admin login")
print("Enter 2 for user login")
print("Enter 3 for exit")
key='user'
key1='password'
while(True):
    a=int(input("Enter the digit to login as admin/user:"))
    if(a==1):
        print("Admin Login")
        username=input("Enter username:")
        if username in lst:
            m1=lst.index(username)
            m1+=1
            password=input("Enter password:")
            if (password==sheet.cell_value(m1,2)):
                print("Login successfull")
                print("Enter 1 for add customer")
                print("Enter 2 for remove customer")
                print("Enter 3 for view details")
                print("Enter 4 for exit")
                while(True):
                    b=int(input("Enter your choice:"))
                    if(b==1):
                        print("Adding the customer\nEnter the customer details")
                        newusername=input("Enter new username:")
                        newpassword=input("Enter new password:")
                        balance=int(input("Enter the balance:"))
                        wt=copy(w1)
                        sheet1=wt.get_sheet(0)
                        sheet1.write(sheet2.nrows,0,newusername)
                        sheet1.write(sheet2.nrows,2,newpassword)
                        sheet1.write(sheet2.nrows,4,balance)
                        wt.save(path1)
                    elif(b==2):
                        print("Remove the customer")
                        lst1=[sheet2.cell_value(k,0) for k in range(1,sheet2.nrows)]
                        while(True):
                            user1=input("Enter the username you want to delete:")
                            if user1 in lst1:
                                m2=lst1.index(user1)
                                m2+=1
                                wt=copy(w1)
                                sheet1=wt.get_sheet(0)
                                sheet1.write(m2,0,'')
                                sheet1.write(m2,2,'')
                                sheet1.write(m2,4,'')
                                
                                wt.save(path1)
                                print("Successfully deleted...")
                                break
                            else:
                                print("Check the username...")
                    elif(b==3):
                        print("View Details")
                        for i in range(sheet2.nrows):
                            for j in range(sheet2.ncols):
                                print(sheet2.cell_value(i,j),end="\t")
                            print()
                    elif(b==4):
                        break
                print("Logged out")
            else:
                print("Enter correct password...")
        else:
            print("Enter correct username")
    elif(a==2):
        print("User Login")
        user_list=[sheet2.cell_value(k,0) for k in range(1,sheet2.nrows)]
        user_name=input("Enter the username of customer:")
        if user_name in user_list:
            user_password=input("Enter the password of customer:")
            m3=user_list.index(user_name)
            m3+=1
            if(user_password==sheet2.cell_value(m3,2)):
                print("Login successfull as user")
                print("Enter 1 for Cash Deposit")
                print("Enter 2 for Cash Withdrawal")
                print("Enter 3 for Password change")
                print("Enter 4 for Balance enquiry")
                print("Enter 5 for logout")
                while(True):
                    user_op=int(input("Enter the digit for respective operation:"))
                    if(user_op==1):
                        print("Cash Deposit...")
                        amount_to_deposit=int(input("Enter the amount to deposit:"))
                        wt=copy(w1)
                        sheet1=wt.get_sheet(0)
                        sheet1.write(m3,4,sheet2.cell_value(m3,4)+amount_to_deposit)
                        print("Amount Deposited Successfully")
                        wt.save(path1)
                    elif(user_op==2):
                        while(True):
                            print("Cash Withdrawal...")
                            amount_to_withdraw=int(input("Enter the amount to withdraw:"))
                            balance=sheet2.cell_value(m3,4)
                            if(amount_to_withdraw<=balance):
                                wt=copy(w1)
                                sheet1=wt.get_sheet(0)
                                sheet1.write(m3,4,sheet2.cell_value(m3,4)-amount_to_withdraw)
                                print("Amount withdrawn successfully...")
                                break
                            else:
                                print("Insufficient balance...")
                    elif(user_op==3):
                        print("Change password...")
                        new_password=input("Enter new password:")
                        confirm_password=input("Confirm password:")
                        if(new_password==confirm_password):
                            wt=copy(w1)
                            sheet1=wt.get_sheet(0)
                            sheet1.write(m3,2,new_password)
                            print("Password changed successfully...")
                            wt.save(path1)
                        else:
                            print("Check the password...")
                            continue
                    elif(user_op==4):
                        print("Bank Balance enquiry...")
                        bank_balance=sheet2.cell_value(m3,4)
                        print("Your Account Balance is: ",bank_balance)
                    elif(user_op==5):
                        break
                    else:
                        print("Wrong digit! Enter the right digit....")
            else:
                print("Check the password....")
        else:
            print("Check the username....")
    elif(a==3):
        break
    else:
        print("Wrong digit! Enter the right digit....")





