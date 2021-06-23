"""
File: DevSweetSpots.py
Author: Clint Kline
Date Created: 06-23-2021
Last Modified: ""
Purpose:    
    This file was created to determine the most lucrative states in which
to pursue a career in software development. This is done by first accessing 
2 seperate files.
    The first file ranks each state by the degree of opportunity each presents 
to aspiring devs. Data from https://www.zippia.com/software-developer-jobs/best-states/ 
was used to determine this rank.
    The second file lists each state and ranks them by their cost of living. A rating
is then assigned to each state. This rating is created by adding the states cost-of-living
rank to its opportunity rank. The state with the lowest rank, being the state with the
most profit potential for devs. 
    This new rating is then written to a new .xlsx file, which is then placed in the current 
working directory where it can be viewed in MS Excel and sorted as desired. 
"""


# need to install pylightxl module with pip before file will execute successfully
import pylightxl as pyxl
import os

# discover location of DevSweetSpots.py and assoc. files
folder = os.path.dirname(__file__)

# assign stateByOpportunity file to variable
# data source: https://www.zippia.com/software-developer-jobs/best-states/
stateByOp = folder + "\\statesByOpportunity.xlsx"
# assign stateByCost file to variable
# data source: https://meric.mo.gov/data/cost-living-data-series
stateByCost = folder + "\\statesByCost.xlsx"

# open stateByOp
sbo = pyxl.readxl(stateByOp)
# open stateByCost
sbc = pyxl.readxl(stateByCost)

sboC1 = sbo.ws(ws='Sheet1').col(col=1)  # represents column 1 in sbo
sbcC1 = sbc.ws(ws='Sheet1').col(col=1)  # represents column 1 in sbc
ratingList = []  # list to append each new rating # to
stateList = []  # new list of states to ensure states pair with correct rating #'s
stateNum = 1  # counter

# for cell in sbo column 1
#   if counter is less than total number of states
#       ID cell contents
#       if state in sbo column 1 is also present in sbc column 1
#           add value of sbo col 2 to sbc col 2, assign to new variable 'rating'
#           append 'rating' to new list
# write new list to new excel file

for cell in sboC1:
    if stateNum <= 51:
        stateName = sbo.ws(ws='Sheet1').address(
            address=('A' + str(stateNum)))
        if stateName in sbcC1:
            rating = sbo.ws(ws='Sheet1').address(address=('B' + str(stateNum))) + \
                sbc.ws(ws='Sheet1').address(address=('B' + str(stateNum)))
            # print(stateName, ":", rating)
        stateList.append(stateName)
        ratingList.append(rating)
        stateNum += 1

# create a blank excel file
newDb = pyxl.Database()
# add a blank worksheet to newDb
newDb.add_ws(ws='Sheet1')

# add stateList to col 1 of newDb
for row_id, data in enumerate(stateList, start=1):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=1, val=data)
# add ratingList to col 2 of newDb
for row_id, data in enumerate(ratingList, start=1):
    newDb.ws(ws='Sheet1').update_index(row=row_id, col=2, val=data)

# write the new file to disk
try:
    pyxl.writexl(
        db=newDb, fn=folder + "\\DevSweetSpotsResults.xlsx")
    print("\nNew spreadsheet successfully created!\n")
except:
    print("File write has Failed.")
