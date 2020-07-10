import openpyxl
import pandas as pd


wb = openpyxl.load_workbook(r'C:\Users\tbarten\Desktop\Projects\Cost_Productivity\PyRun\2020 EFCO Productivity.xlsm', data_only = True)
"""
To use with a different file, change the file path in the above assignment for wb. Use r' to read the file.
Using data_only allows for the script to pull the calculated value in Excel, without it it will show the underlining formula
Commented out print(wb.get_sheet_names()) to only use when I need to check the sheet names for the next step
"""
#print(wb.get_sheet_names()) 


#selecting the tab from the excel file that I want to work in; tab name is a string; keep it in ('')
sheet = wb.get_sheet_by_name('Actual Drivers')
#the tables in the sheets do not align as well as I thought they would, should make a copy of this script and change the sheet and values to match

"""myDict is set to a blank dictionary in prep for the series of list that the next section will add.

Using the below outer for and inner for loops you can work through an array of cells in Excel to get their values.
The outer for loop (i) sets the number of rows to get the values from.
The inner for loop (n) sets the number of columns to get the values from.
With each for loop the first integer in the range is the cell number to start with, it must start with 1 or higher (zero doesn't equal one).
The second integer in the range for the loops is the number of rows or columns that are to be looped through."""

table_dict = \
{"Reissues":{"first":6, "last":11, "uom":"dollars", "group_type":"type"}, \
 "Yield_Pounds":{"first":16, "last":22, "uom":"pounds", "group_type":"department"}, \
 "Window_Units":{"first":61, "last":70, "uom":"units", "group_type":"department"}, \
 "Warehouse_Pieces":{"first":95, "last":96, "uom":"sticks", "group_type":"department"}}



for table in table_dict:
    myDict = {}
    for i in range (5, 6): #for the fixed row of dates
            for n in range (2, 74): 
                data = ([])
                data.append(sheet.cell(row = i, column = n).value)
                myDict.setdefault(n, []).append(data)

        
    for i in range (table_dict[table]['first'], table_dict[table]['last']):
        for n in range (2, 74): 
            data = ([])
            data.append(sheet.cell(row = i, column = n).value)
            myDict.setdefault(n, []).append(data)

    delete = (3,4,5,6,7)

    for d in delete:
        del myDict[d]
    
    """Pandas (pd) are easier to work with when reshaping data for analysis.
    I wanted to use the pandas dataframe to wrangle the data."""    
    df = pd.DataFrame.from_dict(myDict)
    
    """"Since I created the dataframe out of list of list I was stuck with square brackets
    This for loop gets the string of list, effectively removing the square brackets from the dataframe."""
    for r in df:
        df[r] = df[r].str.get(0)
    
 
    """I am sure there is a much cleaner way to reshape a long data set that has a column and row that need to be indexed.
    Either way, I modified the data into a long form with stack, then I unstacked what was the first row, then I used those
    two columns to 'melt' the data into the shape that I was trying to work with."""     
    df=df.set_index([2]).stack().unstack([0]).reset_index()
    df=df.rename(columns={None:'Dates'})
    df=df.melt(id_vars = ['index', 'Dates'])
    
    """Added in some dynamic column name changes. 
    This helps drill with the two common variables of this dataset."""
    group = table_dict[table]['group_type']
    uom = table_dict[table]['uom']
    
    df=df.rename(columns={2:group})
    df=df.rename(columns={'value':uom})
    
    """The data set had this reoccurring month to date (MTD) 
    and year to date (YTD) calculation in it that was not needed. 
    The doesn't equal removes or filters out that calculation."""
    df=df[df.Dates != 'MTD']
    df=df[df.Dates != 'YTD']
    
    df.fillna(0) #making null values zero
    
    #print(df.info())
    
    """save the reshaped dataframe and changed the name to dynamically align with dict key"""
    df.to_csv(r'C:\Users\tbarten\Desktop\Projects\Cost_Productivity\PyRun\actual {0}2020.csv'.format(table))
