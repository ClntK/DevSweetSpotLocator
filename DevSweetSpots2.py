"""
File: DevSweetSpots2.py
Author: Clint Kline
Date Created: 06-23-2021
Last Modified: 06-24-2021
Purpose:
    - This file was created to determine the most lucrative states in which to pursue 
a career in software development. This is done by first accessing 2 seperate files.

    - The first file ranks each state by the degree of opportunity each presents 
to aspiring devs. Data from https://www.zippia.com/software-developer-jobs/best-states/ 
was used to determine this rank.

    - The second file lists each state and ranks them by their cost-of-living. Data from 
https://meric.mo.gov/data/cost-living-data-series was used to determine this rank.

    - A rating is then assigned to each state. This rating is created by adding the 
states cost-of-living rank to its opportunity rank. The state with the lowest rating being 
the state with the most profit potential for devs. This new rating is then written to a new 
.xlsx file, which is then placed in the current working directory where it can be viewed 
in MS Excel and sorted as desired. 

UPDATES:

- 06-24-2021
- New in version 2.
    - I've added columns to represent each states overall score, as well as each states opportunity
    and cost-of-living scores from the accompanying .xlsx files. 
    I've aldo added code to save a finalized version of the file in which all table data is 
    sorted using the 'Rating' column as the sort key. 

"""


# pylightxl module required. install with pip to execute script successfully
import pylightxl as pyxl
import win32com.client
import os

################################
# REQUIRED PYTHON MODULES
################################
# pip3 install pylightxl
# pip3 install pypiwin32 => documentation for this module is sparce, but it works for this files purposes. *shrug*

################################
# ASSIGN VARIABLES
################################

# discover location of 'DevSweetSpots.py' and assoc. files
folder = os.path.dirname(__file__)

# assign statesByOpportunity.xlsx file to variable
# data source: https://www.zippia.com/software-developer-jobs/best-states/
stateByOp = folder + "\\statesByOpportunity.xlsx"
# assign statesByCost.xlsx file to variable
# data source: https://meric.mo.gov/data/cost-living-data-series
stateByCost = folder + "\\statesByCost.xlsx"

# open statesByOpportunity.xlsx
sbo = pyxl.readxl(stateByOp)
# open statesByCost.xlsx
sbc = pyxl.readxl(stateByCost)

# column variables
sboC1 = sbo.ws(ws='Sheet1').col(col=1)  # represents column 1 in sbo
sboC2 = sbo.ws(ws='Sheet1').col(col=2)  # represents column 2 in sbo
sbcC1 = sbc.ws(ws='Sheet1').col(col=1)  # represents column 1 in sbc
sbcC2 = sbc.ws(ws='Sheet1').col(col=2)  # represents column 2 in sbc

# list variables
scoreList = []  # list to append each states score to
stateList = []  # new list of states to ensure states pair with correct rating #'s
ratingList = []  # list to append each new rating # to
opList = []  # list of opportunity ratings assigned to relevant state
costList = []  # list of cost-of-living ratings assigned to relevant state

# counter
stateNum = 1

################################
# DATA COLLECTION/CREATION
################################

for cell in sboC1:  # for cell in sbo column 1
    if stateNum <= 51:  # if counter is less than total number of states
        stateName = sbo.ws(ws='Sheet1').address(
            address=('A' + str(stateNum)))  # Identify cell contents of state column in 'statesByOpportunity.xlsx'
        if stateName in sbcC1:  # if state in sbo column 1 is also present in sbc column 1
            rating = sbo.ws(ws='Sheet1').address(address=('B' + str(stateNum))) + \
                sbc.ws(ws='Sheet1').address(address=('B' + str(stateNum))
                                            )  # add value of sbo col 2 to sbc col 2, assign to new variable 'rating'
        scoreList.append(stateNum)
        stateList.append(stateName)  # append state name to new list
        ratingList.append(rating)  # append 'rating' to new list
        opList.append(sbo.ws(ws='Sheet1').address(
            address=('B' + str(stateNum))))  # include Op score
        costList.append(sbc.ws(ws='Sheet1').address(
            address=('B' + str(stateNum))))  # include cost-of-living score
        stateNum += 1  # add 1 to row counter

################################
# INITIATE SPREADSHEET CREATION
################################

# create a blank excel file
newDb = pyxl.Database()
# add a blank worksheet to newDb
newDb.add_ws(ws='Sheet1')

################################
# CREATE HEADERS
################################

newDb.ws(ws='Sheet1').update_index(row=1, col=1, val='Score')
newDb.ws(ws='Sheet1').update_index(row=1, col=2, val='State')
newDb.ws(ws='Sheet1').update_index(row=1, col=3, val='Rating')
newDb.ws(ws='Sheet1').update_index(row=1, col=4, val='Op Score')
newDb.ws(ws='Sheet1').update_index(row=1, col=5, val='Cost Score')

################################
# POPULATE DATABASE
################################

# create numbered index in column 1
for row_id, data in enumerate(scoreList, start=2):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=1, val=data)
# add stateList to col 2 of newDb
for row_id, data in enumerate(stateList, start=2):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=2, val=data)
# add ratingList to col 3 of newDb
for row_id, data in enumerate(ratingList, start=2):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=3, val=data)
# include opportunity score in column 4
for row_id, data in enumerate(opList, start=2):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=4, val=data)
# include cost-of-living score in column 5
for row_id, data in enumerate(costList, start=2):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=5, val=data)


################################
# SAVE OUTPUT
################################

try:
    # write the new file to disk
    pyxl.writexl(
        db=newDb, fn=folder + "\\DevSweetSpotsResults.xlsx")

    ################################
    # SORT OUTPUT
    ################################
    # sort table data by rating
    # assign Excel app to variable
    excel = win32com.client.Dispatch("Excel.Application")
    # designate file and workbook on which to perform sort function
    file = excel.Workbooks.Open(folder + "\\DevSweetSpotsResults.xlsx")
    table = file.Worksheets('Sheet1')
    # designate cell to sort. B2=start sort @ column be row 2. BE52=sort columns B-E, stop at row 52. C2=sort key
    table.Range('B2:BE52').Sort(Key1=table.Range(
        'C2'), Order1=1, Orientation=1)
    # finish up by saving and closing file
    file.Save()
    excel.Application.Quit()

    print("\nNew spreadsheet successfully created!\n")
except Exception as e:
    print("\nFile write has Failed.\n", "Reason: ", e)
