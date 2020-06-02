import openpyxl
import pandas as pd


wb = openpyxl.load_workbook(r'filepath.xlsm', data_only = True)
"""
To use with a different file, change the file path in the above assignment for wb. Use r' to read the file.
Using data_only allows for the script to pull the calculated value in Excel, without it it will show the underlining formula
Commented out print(wb.get_sheet_names()) to only use when I need to check the sheet names for the next step
"""
#print(wb.get_sheet_names()) 


#selecting the tab from the excel file that I want to work in; tab name is a string; keep it in ('')
sheet = wb.get_sheet_by_name('worksheet_name')

myDict = {}

"""
myDict is set to a blank dictionary in prep for the series of list that the next section will add.

Using the below outer for and inner for loops you can work through an array of cells in Excel to get their values.
The outer for loop (i) sets the number of rows to get the values from.
The inner for loop (n) sets the number of columns to get the values from.
With each for loop the first integer in the range is the cell number to start with, it must start with 1 or higher (zero doesn't equal one).
The second integer in the range for the loops is the number of rows or columns that are to be looped through.
Note: if the first digit in the column loop (i.e. currently it is a 2) the index will also need to be changed.
"""

for i in range (5, 6): #for the fixed row of dates at the top of the sheet
    for n in range (2, 74): #for the columns and width of the table
        data = ([])
        data.append(sheet.cell(row = i, column = n).value)
        myDict.setdefault(n, []).append(data)

        
for i in range (6, 10): #change for different rows of data
    for n in range (2, 74): #must be the same as the loop above
        data = ([])
        data.append(sheet.cell(row = i, column = n).value)
        myDict.setdefault(n, []).append(data)

delete = (3,4,5,6,7)

for d in delete:
    del myDict[d]
""" 
Removing some unwanted noise from the dataset.
Data contains calculations for month end and year end. 
The 'delete' assignment was made instead of looping through the range incase other columns need to be added for other sheets.
"""

df = pd.DataFrame.from_dict(myDict)

"""
Pandas (pd) are easier to work with when reshaping data for analysis.
I wanted to use the pandas dataframe to wrangle the data.
"""

for r in df:
    df[r] = df[r].str.get(0)
""""
Since I created the dataframe out of list of list I was stuck with square brackets
This for loop gets the string of list, effectively removing the square brackets from the dataframe.
"""


df=df.set_index([2]).stack().unstack([0]).reset_index()
df=df.rename(columns={None:'Dates'})
df=df.melt(id_vars = ['index', 'Dates'])
df=df.rename(columns={2:'RI_Type'})
"""
I am sure there is a much cleaner way to reshape a long data set that has a column and row that need to be indexed.
Either way, I modified the data into a long form with stack, then I unstacked what was the first row, then I used those
two columns to 'melt' the data into the shape that I was trying to work with.
""" 

df=df[df.Dates != 'MTD']
df=df[df.Dates != 'YTD']
"""
The data set had this reoccurring month to date (MTD) 
and year to date (YTD) calculation in it that was not needed. 
The doesn't equal removes or filters out that calculation.
"""

print(df.head())

df.to_csv(r'newfilepath.csv')

"""
I like to display the top rows of the of the dataframe (i.e. print(df.head())) so that I have a heads up
of what the data looks like before it is sent to the file. It won't change the file or stop the process if it isn't correct. 
The df.to_csv then writes it to the sheet that I want to work with. For now, I will send it to a new sheet each time
that I run this code. I will also change the initial for range with i & n to grab the multiple tables form this sheet. 
"""
