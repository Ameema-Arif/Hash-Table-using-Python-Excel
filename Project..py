import os
import openpyxl
from tabulate import tabulate

os.chdir('d:\\DSAProject')

#Accessing the excel worksheet contaning the hash table

wb = openpyxl.load_workbook('HashTable.xlsx')
Directory = wb.get_sheet_by_name('Directory')


choice =  int(input("Press 1 to add a new record\nPress 2 to search a record\n:"))

#We've to enter atleast 30 records in the table that's why
n = 30
#We are taking total 75 records becuase there are total 75 employees so in order to minimize the no. of collisions we will take any prime no. lesser than 75
m = 73

#for adding a new record
if choice == 1:

#taking while loop ,so we can perform the same operation several times if needed
    while m>n:
        Enum = int(input("Enter a 5-digit Employee no. containing integers only:"))
        if Enum>9999 and Enum<100000:
            Ename = str(input("Enter Employee name(Don't leave this empty):"))
            Landline = int(input("Enter Landline no.(e.g:021*******)(Don't leave this empty):"))
            Mobile = int(input("Enter Mobile no.(e.g:03*********)(Don't leave this empty):"))

            Hashvalue = int(Enum%73)
            if Directory['A'+str(Hashvalue)].value == None :
                if Directory['B'+str(Hashvalue)].value == None :
                    if Directory['C'+str(Hashvalue)].value == None:
                        
                        Directory['A'+str(Hashvalue)].value = Enum
                        Directory['A'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

                        Directory['B'+str(Hashvalue)].value = Ename
                        Directory['B'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

                        Directory['C'+str(Hashvalue)].value = Landline
                        Directory['C'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        Directory['D'+str(Hashvalue)].value = Mobile
                        Directory['D'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        wb.save('HashTable.xlsx')
                        print("Record added successfully!")

                        select = int(input("Press 1 to add another record\nPress 2 to Quit\n:"))
                        if select == 1:
                            continue
                        elif select == 2:
                            break


            elif Directory['A'+str(Hashvalue)].value != None :
                if Directory['B'+str(Hashvalue)].value != None :
                    if Directory['C'+str(Hashvalue)].value != None:
                        while Directory['A'+str(Hashvalue)].value != None and Directory['B'+str(Hashvalue)].value != None and Directory['C'+str(Hashvalue)].value != None :
                            if Hashvalue < 75:
                                Hashvalue = Hashvalue+1

                            elif Hashvalue == 75:
                                Hashvalue = 1

                        Directory['A'+str(Hashvalue)].value = Enum
                        Directory['A'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        Directory['B'+str(Hashvalue)].value = Ename
                        Directory['B'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        Directory['C'+str(Hashvalue)].value = Landline
                        Directory['C'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        Directory['D'+str(Hashvalue)].value = Mobile
                        Directory['D'+str(Hashvalue)].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                        
                        wb.save('HashTable.xlsx')
                        print("Record added successfully!")

                        select = int(input("Press 1 to add another record\nPress 2 to Quit\n:"))
                        if select == 1:
                            continue
                        elif select == 2:
                            break

                        
        else:
            print("Invalid Employee no.!")
            continue



#for Searching a record
elif choice == 2 :

#taking while loop ,so we can perform the same operation several times if needed
    while m>n:
        Enum = int(input("Enter a 5-digit Employee no. containing integers only:"))

        if Enum>9999 and Enum<100000:
            Hashvalue = int(Enum%73)
            
            if Directory['A'+str(Hashvalue)].value == Enum:
                num = Directory['A'+str(Hashvalue)].value
                name = Directory['B'+str(Hashvalue)].value
                phone = Directory['C'+str(Hashvalue)].value
                cell = Directory['D'+str(Hashvalue)].value

                Data = [[str(num),str(name),str(phone),str(cell)]]
                Headers = ["Employee No.","Employee Name","Landline No.","Mobile No."]

                print(tabulate(Data,headers=Headers))

                select = int(input("\nPress 1 to add another record\nPress 2 to Quit\n:"))
                if select == 1:
                    continue
                elif select == 2:
                    break



            elif Directory['A'+str(Hashvalue)].value != Enum:
                while Directory['A'+str(Hashvalue)].value != Enum:
                    if Hashvalue < 75:
                        Hashvalue=Hashvalue+1

                    elif Hashvalue==75:
                        Hashvalue=1

                
                num = Directory['A'+str(Hashvalue)].value
                name = Directory['B'+str(Hashvalue)].value
                phone = Directory['C'+str(Hashvalue)].value
                cell = Directory['D'+str(Hashvalue)].value

                Data = [(str(num),str(name),str(phone),str(cell))]
                Headers = ["Employee No.","Employee Name","Landline No.","Mobile No."]

                print(tabulate(Data,headers=Headers))

                select = int(input("\nPress 1 to add another record\nPress 2 to Quit\n:"))
                if select == 1:
                    continue
                elif select == 2:
                    break


        else:
            print("Invalid Employee no.!")
            continue

        
               























                     
