




# Import needed modules for this script; Pandas for dataframes, datetime for "Today" variable, xlsxwriter for the export to excel section, numpy for adding a null column, os because it is awesome and I felt like it

import pandas as pd
import os
import datetime 
import xlsxwriter
import numpy as np

# Set variable "Today" to determine how many of each part should be added to cycle count sheets.

Now = datetime.datetime.now()
Today = Now.weekday()

# Create Base dataframes to import data and then create lists.  Part and Location Data should be pulled from Fishbowl data queries "Cycle Count Parts (redoux)" and "Cycle Count Location"
# Cycle count dataframe is where the parts to be counted are stored

cyclecount = pd.DataFrame(data=None, index=None, columns = ['PartNumber', 'PartDescription', 'Age','Qty'],dtype=None,copy=False)
part = pd.read_excel('Z:\\Cycle Counting\\Daily List\\raw data\\Part.xlsx')
Location = pd.read_excel('Z:\\Cycle Counting\\Daily List\\raw data\\Locations.xlsx')

# All parts have a counter date that is reset when cycle counts.  To avoid having many days where nothing needs to be counted, this part of the script removes any parts that are not in inventory from part list.  This will not reset the counter though.
#The Location dataframe is merged onto the part dataframe, any parts with no locations active are dropped from the DataFrame.  Finally age is determined by subtracting the date of last counts from today

part = pd.merge(part,Location.copy(),how='left',on='PartNumber')
part = part.dropna()
Part = part.drop(['LocGroup','Location'],1)
part['Age'] = part['today'] - part['CF-Cycle Count Date']

# Separate A-Parts, find most needed counts, move relevant parts to cyclecount sheet based on day of week
# A_Part Dataframe is created by pulling all parts from the part dataframe with the ABCcode of 'A'.  The A_Part DF is then sorted by age with oldest on top
# A_Part_Even and A_Part_Odd give the correct number of parts to be appended to the cycle count DF based on the day of the week.
# The "IF" statement determines the day of the week and then appends the correct df onto the cyclecount DF

A_Part = part[['PartNumber','PartDescription','Age']][part['ABCCode'] == 'A'].copy()
A_Part = A_Part.sort_values(by='Age',axis=0,ascending=False,)
A_Part_Even = A_Part.head(3)
A_Part_Odd = A_Part.head(2)
if Today == 0:
        cyclecount = cyclecount.append(A_Part_Even,ignore_index=True,verify_integrity=False)
elif Today == 1:
        cyclecount = cyclecount.append(A_Part_Odd,ignore_index=True,verify_integrity=False)
elif Today == 2:
        cyclecount = cyclecount.append(A_Part_Even,ignore_index=True,verify_integrity=False)
elif Today == 3:
        cyclecount = cyclecount.append(A_Part_Odd,ignore_index=True,verify_integrity=False)
else:
        cyclecount = cyclecount.append(A_Part_Even,ignore_index=True,verify_integrity=False)

# Separate B-Parts, find most needed counts, move relevant parts to cyclecount sheet.  B parts are always 1 per day so no "IF" statement is needed
# B_Part DF is created by pulling all parts from the part df with an ABCcode of B.  The B_Part DF is then sorted by age with oldest on top
# B_Part_All takes the top row of the B_Part df and then appends it to cyclecount df
B_Part = part[['PartNumber','PartDescription','Age']][part['ABCCode'] == 'B'].copy()
B_Part = B_Part.sort_values(by='Age',axis=0,ascending=False,)
B_Part_All = B_Part.head(1)
cyclecount = cyclecount.append(B_Part_All)

# Separate C-Parts, find most needed counts, move relevant parts to cyclecount sheet based on day of week
# C_Part Dataframe is created by pulling all parts from the part dataframe with the ABCcode of 'C'.  The C_Part DF is then sorted by age with oldest on top
# C_Part_All give the correct number of parts to be appended to the cycle count DF based on the day of the week.
# The "IF" statement determines the day of the week and then appends the correct df onto the cyclecount DF

C_Part = part[['PartNumber','PartDescription','Age']][part['ABCCode'] == 'C'].copy()
C_Part = C_Part.sort_values(by='Age',axis=0,ascending=False,)
C_Part_All = C_Part.head(1)
if Today == 1:
        cyclecount = cyclecount.append(C_Part_All,ignore_index=True,verify_integrity=False)
elif Today == 3:
        cyclecount = cyclecount.append(C_Part_All,ignore_index=True,verify_integrity=False)

# Count the number of parts on the cycle count,
#Rows = cyclecount.count()

# Merge locations onto cyclecount dataframe create new dataframe called export
# Sort the export dataframe by "LocGroup" and then drop the "Age" column
# add a blank column titled counter

export = pd.merge(cyclecount.copy(),Location.copy(),how='left',on='PartNumber')
export = export.sort_values(by='LocGroup',axis=0,ascending=False)
export = export.drop(['Age'],1)
export['Counter'] = np.nan

# Create individual sheets for each Warehouse (All non manned buildings are lumped into WA)
# Adds all lines with the "LocGroup" for each remote warehouse to a new export, then drops those lines from the export df

export_SAC = export[['Counter', 'PartNumber', 'PartDescription', 'Qty','LocGroup','Location']][export['LocGroup']== 'SAC'].copy()
export = export[export.LocGroup != 'SAC']
export_MD = export[['Counter', 'PartNumber', 'PartDescription', 'Qty','LocGroup','Location']][export['LocGroup']== 'MD'].copy()
export = export[export.LocGroup != 'MD']
export_TN = export[['Counter', 'PartNumber', 'PartDescription', 'Qty','LocGroup','Location']][export['LocGroup']== 'TN'].copy()
export = export[export.LocGroup != 'TN']
export_WA = export

#Write excel files for each of the location groups then saves them to the Shared Services drive on the NAS

writer = pd.ExcelWriter('Z:\Cycle Counting\Daily List\Daily List.xlsx',engine='xlsxwriter')
export_MD.to_excel(writer,sheet_name='MD',index=False)
export_SAC.to_excel(writer,sheet_name='SAC',index=False)
export_TN.to_excel(writer,sheet_name='TN',index=False)
export_WA.to_excel(writer,sheet_name='WA',index=False)
writer.save()

