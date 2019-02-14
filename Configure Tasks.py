"""Iterate through Excel file and find: the frequency of specific names, the number of times each of them corresponded with the word "Completed" and subtract their corresponding dates from each other"""
"""This is to determine how many tasks each individual was assigned and how many of them have they completed or missed"""
""" 1. Number of assigned tasks = the amount of times their name was mentioned in the "Assigned To" column in the spreadsheet """
""" 2. Number of completed tasks = the amount of times their name was mentioned in the "Assigned To" column with the word "Completed" in the "Status" Column """
""" 3. Number of overdue tasks = amount of times their name was mentioned in the "Assigned To" column with "End date" - today's date + "Status" = " Active " """
from __future__ import print_function
import pandas
import xlrd
import re
import string
import datetime
from collections import Counter
from astropy.table import Table, Column
from tabulate import tabulate

#Tasks Assigned per person """Tasks Assigned"""-------------------------------------------------------------------------------------
#Read the Excel file
df = pandas.read_excel('excel_file.xls')

#Assignees column in Excel
assignees_raw = df["Assigned To"]

#Assignees to be counted
Names = ["Name A", "Name B", "Name C", "Name D", "Name E", "Name F", "Name G", "Name H", "Name I", "Name J"]

#List of Trues for each Name
True_lists = []
for name in Names:
    True_list = assignees_raw.str.contains(name, case=True).tolist()
    True_lists.append(True_list)

#Counter for each Assignee
#Mk, Ng, Bl, Ec, Wh, Cs, Sl, Dt, Dw, As
Tasks_Assigned = []
def count_assigned_tasks(Lists):
    for l in range(len(Lists)):
        counter = 0
        for bol in Lists[l]:
            if bol == True:
                counter += 1
        Tasks_Assigned.append(counter)

count_assigned_tasks(True_lists)

#Tasks Completed per person """Tasks Completed"""------------------------------------------------------------------------------------
#Status column in Excel
Tasks_Completed = []
status_raw = df["Status"]

#Status Range = True list Range
#Name = Counter Range = True lists Range
def count_completed_tasks(Lists):
    for List in Lists:
        counter = 0
        for i in range(len(List)):
            if (status_raw[i] == "Completed") and (List[i] == True):
                counter += 1
        Tasks_Completed.append(counter)

count_completed_tasks(True_lists)
#print Tasks_Completed 

#Tasks Overdue per person """Tasks Overdue"""------------------------------------------------------------------------------------
#Overdue column in Excel
Tasks_Overdue = []
Start_Date = df["Start Date"]
End_Date = df["End Date"]
Current_Date = datetime.datetime.now()

def count_overdue(Lists):
    for List in Lists:
        counter = 0
        for i in range(len(List)):
            if (status_raw[i] == "Active") and (List[i] == True) and (Current_Date > End_Date[i]):
                counter += 1
        Tasks_Overdue.append(counter)

count_overdue(True_lists)
#print Tasks_Overdue 

# Configure Final Result
table = Table([Names, Tasks_Assigned, Tasks_Completed, Tasks_Overdue], names=("Names", "Tasks Assigned", "Tasks Completed", "Tasks Overdue"))
print (table)

#Write the results table to text file
with open('Result.txt', 'w') as outputfile:
    print(table, file=outputfile)

