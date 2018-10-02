import os
import glob
import csv
from xlsxwriter.workbook import Workbook


#Declaring Variables
row_num = int()
col_num = int()
col_name = []
data = list()
fileName = str()

# """""""""""""""""""""""""""""
# Welcome text and instructions
# """""""""""""""""""""""""""""
def welcome():
    print("/=======================================================/")
    print("WELCOME TO EXEL GENERATOR")
    print("Never include DATE and Serial Number in Row and Colum")
    print("it ll generate it automatically")
    print("/=======================================================/")

#""""""""""""""""""""""""""""""
# Function for Asking question
# (y/n)
#""""""""""""""""""""""""""""""
def ask(question):
    print(question + " (y/n): ")
    reply = str(input().lower().strip())
    if reply[0] == 'y':
        return True
    if reply[0] == 'n':
        return False
    else:
        return ask("... please enter ")

#""""""""""""""""""""""""""""""""""""
# function for asking column Names
# not including date and serial number
#""""""""""""""""""""""""""""""""""""
def ask_col_names():
    print("Enter {} Colum Names: ".format(col_num))
    for i in range(col_num):
        col_name.append(str(input()))



#""""""""""""""""""""""""""""""""""""""
# function for taking data input
# in data 2d list
#""""""""""""""""""""""""""""""""""""""
def take_data_input():
     for col in range(col_num):
         for row in range(row_num):
             print("{} {} :".format(col_name[col] ,row+1) , end="")
             data[col][row] = input()



#""""""""""""""""""""""""""""""""""""""""
# Function for Asking for
# Serial Numbers and adding functionility
#""""""""""""""""""""""""""""""""""""""""
def ask_sn():
    if ask("Do you want Serial Numbers?"):
        serial_num = [str(x+1) for x in range(row_num)]
        col_name.insert(0,'S.No')
        data.insert(0, serial_num) #inserting serial number list in Data list
        global col_num
        col_num = col_num + 1

#"""""""""""""""""""""""""
# Function for Asking date
# and adding functionality
#"""""""""""""""""""""""""
def ask_date():
    if ask("Yes I want auto date insertion | No I entered it manually"):
        date = []
        print("Enter year:")
        year = str(input())
        print("Enter Month:")
        month = str(input())
        for i in range(0,row_num):
            print("Enter day of Row {}:".format(i+1))
            day = str(input())
            date.append(day + '.' + month + '.' + year)

        col_name.insert(1,'Date')
        data.insert(1, date) #inserting date list in Data list
        global col_num
        col_num = col_num + 1
        return date


#""""""""""""""""""""""""""""""""""""""
# function for writing Column names
# in CSV file
#""""""""""""""""""""""""""""""""""""""
def write_column_name_in_csv():
    #Writing Row in CSV File
    with open(fileName , 'w+') as sheet:
        for i in range(col_num):
            sheet.write(str(col_name[i]) + ",")

        sheet.write("\n")


#""""""""""""""""""""""""""""""""""""""
# very important function for writing
# Data in CSV file (where magic happens)
#""""""""""""""""""""""""""""""""""""""
def write_data_in_csv():
    #Writing Row in CSV File
    with open(fileName , "a") as sheet:
        for row in range(row_num):
            for col in range(col_num):
                sheet.write(str(data[col][row])+",")

            sheet.write("\n")

    print("\n Sheet Written!!")

def csv_to_exel():
    for csvfile in glob.glob(os.path.join('.', 'data.csv')):
        workbook = Workbook(csvfile[:-4] + '.xlsx')
        worksheet = workbook.add_worksheet()
        with open(csvfile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()



if __name__ == '__main__':
        welcome()
        fileName = str(input("Enter name of File"))
        row_num = int(input("Input number of rows: "))
        col_num = int(input("Input number of columns: "))
        data = [[0 for col in range(row_num)] for col in range(col_num)]
        ask_col_names() # asking column names without sn and date
        take_data_input() #taking data input
        print(data)
        ask_sn() #asking sn
        ask_date() #asking date
        print(data)
        write_column_name_in_csv()
        write_data_in_csv()
        csv_to_exel()
        os.remove("data.csv")
