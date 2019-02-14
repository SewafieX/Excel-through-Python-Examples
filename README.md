# Excel-through-Python-Examples
Iterate through Excel file and find: the frequency of specific names, the number of times each of them corresponded with the word "Completed" and subtract their corresponding dates from each other.
This is to determine how many tasks each individual was assigned and how many of them have they completed or missed
  1. Number of assigned tasks = the amount of times their name was mentioned in the "Assigned To" column in the spreadsheet 
  2. Number of completed tasks = the amount of times their name was mentioned in the "Assigned To" column with the word "Completed" in the   "Status" Column
  3. Number of overdue tasks = amount of times their name was mentioned in the "Assigned To" column with "End date" - today's date +         "Status" = " Active " 
