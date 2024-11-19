# Data Prep for TPA data dump from csv
### Helper Tab
1. You will need an excel file with a helper tab that contains the following named cells and formulas. This allows the csv file to be named as indicated and dropped into a folder with the excel worksheet where the Power Query code will be executed.
   FilePath=SUBSTITUTE(LEFT(CELL("filename",B2),SEARCH("]",CELL("filename",B2))-1),"[","")
   FilePathFolder=SUBSTITUTE(FilePath,RIGHT(FilePath,LEN(FilePath)-FIND("@",SUBSTITUTE(FilePath,"\","@",LEN(FilePath)-LEN(SUBSTITUTE(FilePath,"\",""))),1)),"")
   FileName=broadspire.xlsx
2. A table needs to be created with three columns.
  Effective=Effective date of policy term
  Deductible=Deductible amount in dollars
  LOB=Coverage type (eg. AL, CGL, WC)

### PQ Files
You will need all .pq files in this repor, as there are some dependencies amonst them. Working together, they identify multiple features for the same claims, aggregate the financial data, retaining mulitple claim numbers and claimant names, as needed, and output a single coverage line on each of three tabs. 
  
   
