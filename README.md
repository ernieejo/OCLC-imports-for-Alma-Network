# OCLC-imports-for-Alma-Network
OCLC BIB processing report – Script Instructions 

Use these scripts to prepare spreadsheets that can be reviewed to create files for updating and maintaining OCLC numbers in Alma. Prior to running these scripts, some one-time setup is needed.  

* You need to create a base folder where the scripts and the files they create will be stored.

* Add a blank Excel file called "Do_not_change.xlsx" to the base folder. In reviewing the OCLC numbers, there are situations where its best to ignore certain records. The do not change file can be used to list MMS IDs for records that you want the script to ignore.

* You need to change the file paths within Script 1 and Script 2 so that they point to the folder you’ve created. Leave the file name that they point to within that folder as is.  

* Create an Alma Analytics with the following subjects in this order: MMS ID, OCLC Control Number (035a), OCLC Control Number (035z), Network ID. Save the Analytic as "OCLC-doublecheck"

Once you have a base folder and file path established, follow the steps below to create files for NZ processing of OCLC BIB processing report data. It is important that the entire process be done within the shortest timeframe possible. 

<b>Step1.</b> 

* Retrieve the BIB processing report from OCLC WorldShare manager and save to your base folder. Name the file BIB_processing_report.txt 

<b>Step2.</b> 

* Run Script1 - OCLC-IZ Comparison File

This will extract the records that need to be reviewed and updated and create a spreadsheet with two tabs:
1. DIFF - Use the MMS IDs from this sheet to filter the Alma Analytic 
2. SAME - No action is needed on these records

* Export the filtered analytic as an Excel sheet and save in base folder. Name the file "OCLC-doublecheck.xlsx"

<b>Step3.</b> 

* Run Script2 - OCLC-NZ Comparison File

This will create a new Excel file in your base folder "comparison_fileNZ"
It compares the DIFF tab of the OCLC-IZ comparison file to OCLC-doublecheck.xlsx" and "Do_not_change.xlsx"
This provides current data to avoid duplication in 035z
The file comparison_fileNZ can be used to review and update OCLC numbers in Alma.


