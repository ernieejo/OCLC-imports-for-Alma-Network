# OCLC-imports-for-Alma-Network
OCLC BIB processing report – Script Instructions 

Use these scripts to prepare spreadsheets that can be reviewed to create files for updating and maintaining OCLC numbers in Alma. Prior to running these scripts, some one-time setup is needed.  

* You need to create a base folder where the scripts and the files they create will be stored.

* You need to change the file paths in the script so that they point to the folder you’ve created. 

* Create an Alma Analytics with the following subjects in this order: MMS ID, OCLC Control Number (035a), OCLC Control Number (035z), Network ID. Save the Analytic as "OCLC-doublecheck"

Once you have a base folder and file path established, follow the steps below to create a review spreadsheet for the OCLC BIB processing report data. It is important that the entire process be done within the shortest timeframe possible. For example if the institution's records are published to OCLC on a weekly bases, the reports produced should be processed weekly.

<b>Step1.</b> 

* Retrieve the BIB processing report from OCLC WorldShare manager and save to your base folder. Name the file BIB_processing_report.txt 

<b>Step2.</b> 

* Run Script1 - OCLC-IZ Comparison File

This will extract the records that need to be reviewed and updated and create a spreadsheet with two tabs:
1. DIFF - Use the MMS IDs from this sheet to filter the Alma Analytic 
2. SAME - No action is needed on these records

* To grab Network IDs for items, you can use OCLC-doublecheck Alma Analytic. This analytic can also help you to determine if the record has already been updated and where there are existing 035a fields that need to be moved to the 035z.
* Review the records listed in the DIFF tab of the spreadsheet output to determine what changes (if any) should be made to the bibliographic record in Alma.



