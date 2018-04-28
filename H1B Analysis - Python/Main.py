# import pyodbc for connecting to the local database and json for writing objects
import pyodbc
import json
from Model.Company import Company
from Model.Location import Location
from Model.WageGroup import WageGroup
from Model.Wages import Wages

# establish a connection to the mdf file through SQLExpress server
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=DESKTOP-QO716QE\SQLEXPRESS;"
                      "Database=H1BVisa;"
                      "Trusted_Connection=yes;")
# PART 1
# a database cursor is a control structure that enables traversal over the records in a database, 
# retrieval, addition and removal of database records
cursor = cnxn.cursor()
sqlCommand = ("select * from v_CaseStatus_Count order by Certified desc")
cursor.execute(sqlCommand) # execute the read statement
# retrieves the next row of a query result set and returns a single sequence, 
# or None if no more rows are available
results = cursor.fetchone() 

# create an empty list of company objects
companies = []
# while results fetches all rows till none available
while results:
    
    # 1.1 instantiating a company object using the parameterized constructor
    aCompany = Company(results[0], results[1], results[2], results[3], results[4])
    # 1.2 printing out __repr__ string representation of company object
    
    # 1.3 adding each new company object to the list
    companies.append(aCompany)
    # results stores the row returned by the cursor
    results = cursor.fetchone()

# printing the companies list
print(companies)
# printing the number of company objects in the list 
print('Length of companies list is', len(companies))

# 1.4 opening a file named company.json with write access rights
fo = open("company.json", "w+")
# using json dump function to write the students list object into the file in a dictionary format 
json.dump(companies, fo, default=lambda o:o.__dict__)
# close the file
fo.close()


#############################################################################
# PART 2
sqlCommand = ("SELECT * FROM [H1BVisa].[dbo].[AVG_WAGES]")
cursor.execute(sqlCommand) # execute the read statement
# retrieves the next row of a query result set and returns a single sequence, 
# or None if no more rows are available
results = cursor.fetchone() 

# create an empty list of wage objects
wages = []
# while results fetches all rows till none available
while results:
    
    # 2.1 instantiating a wage object using the parameterized constructor
    aWage = Wages(results[0], results[1], results[2], results[3])
    # 2.2 printing out __repr__ string representation of wage object
    
    # 2.3 adding each new wage object to the list
    wages.append(aWage)
    # results stores the row returned by the cursor
    results = cursor.fetchone()

# printing the wages list
print(wages)
# printing the number of wage objects in the list 
print('Length of wages list is', len(wages))

# 2.4 opening a file named wage.json with write access rights
fo = open("wage.json", "w+")
# using json dump function to write the wages list object into the file in a dictionary format 
json.dump(wages, fo, default=lambda o:o.__dict__)
# close the file
fo.close()


#############################################################################
# PART 3
sqlCommand = ("select * from v_Jobs_By_State")
cursor.execute(sqlCommand) # execute the read statement
# retrieves the next row of a query result set and returns a single sequence, 
# or None if no more rows are available
results = cursor.fetchone() 

# create an empty list of locations objects
locations = []
# while results fetches all rows till none available
while results:
    
    # 3.1 instantiating a location object using the parameterized constructor
    aLocation = Location(results[0], results[1], results[2], results[3])
    # 3.2 printing out __repr__ string representation of location object
    
    # 3.3 adding each new locations object to the list
    locations.append(aLocation)
    # results stores the row returned by the cursor
    results = cursor.fetchone()

# printing the locations list
print(locations)
# printing the number of location objects in the list 
print('Length of location list is', len(locations))

# 3.4 opening a file named location.json with write access rights
fo = open("location.json", "w+")
# using json dump function to write the locations list object into the file in a dictionary format 
json.dump(locations, fo, default=lambda o:o.__dict__)
# close the file
fo.close()

#############################################################################
# PART 4
sqlCommand = ("select * from v_Cases_By_IncomeGroup ORDER BY Certified desc")
cursor.execute(sqlCommand) # execute the read statement
# retrieves the next row of a query result set and returns a single sequence, 
# or None if no more rows are available
results = cursor.fetchone() 

# create an empty list of wagegroup objects
wagegroups = []
# while results fetches all rows till none available
while results:
    
    # 4.1 instantiating a wagegroup object using the parameterized constructor
    aWageGroup = WageGroup(results[0], results[1], results[2], results[3], results[4])
    # 4.2 printing out __repr__ string representation of wagegroup object
    
    # 4.3 adding each new wagegroup object to the list
    wagegroups.append(aWageGroup)
    # results stores the row returned by the cursor
    results = cursor.fetchone()

# printing the wagegroup list
print(wagegroups)
# printing the number of wagegroups objects in the list 
print('Length of wagegroups list is', len(wagegroups))

# 4.4 opening a file named wagegroup.json with write access rights
fo = open("wagegroup.json", "w+")
# using json dump function to write the wagegroups list object into the file in a dictionary format 
json.dump(wagegroups, fo, default=lambda o:o.__dict__)
# close the file
fo.close()


# close the database connection
cnxn.close()