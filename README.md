# OCLC-imports-for-Alma-Network
OCLC BIB processing report – Script Instructions 

Use these scripts to prepare spreadsheets that can be reviewed to create files for updating and maintaining OCLC numbers in Alma. Prior to running these scripts, some one-time setup is needed.  

* You need to create a base folder where the scripts and the files they create will be stored.  

* You need to change the file paths within Script 1 and Script 2 so that they point to the folder you’ve created. Leave the file name that they point to within that folder as is.  

* Create an Alma Analytics with the following subjects in this order: MMS ID, OCLC Control Number (035a), OCLC Control Number (035z), Network ID 

Once you have a base folder and file path established, follow the steps below to create files for NZ processing of OCLC BIB processing report data. It is important that the entire process be done within the shortest timeframe possible. 

<b>Step1.</b> 

* Retrieve the BIB processing report from OCLC WorldShare manager and save to your base folder. Name the file BIB_processing_report.txt 

<b>Step2.</b> 

* Run Script1.  

* Use the MMS IDs in the ‘DIFF’ sheet of the file comparison_fileIZ produced by the script to filter your Alma Analytic 

* Export the filtered analytic and save in base folder. Name the file OCLC-doublecheck.xlsx 

<b>Step3.</b> 

* Review DIFF and a_to_z sheets in the file comparison_fileNZ produced by the scripts. 

* The ‘DIFF’ sheet shows all records where the OCLC Control Number (035a) is different than the 035 $a 

    * Use it to locate and correct any records with two OCLC numbers in the 035a (filter using contains”;”) 

* The a_to_z sheet shows all the records where the OCLC Control Number (035a) is different than the OCLC Control Number (035z) 

    * Review for instances where there are multiple 035z in the record (filter using contains”;”). These need to be manually removed from the list one of the OCLC Control Number (035z) match the OCLC Control Number (035a). 

<b>Step4.</b>

* Create two files to send to NZ for processing. 

1. Create a text file with the header MMS ID that contains the column of Network ID numbers from the a_to_z tab. Name the file InstitutionName_a_to_z.txt 

    * This file will be used to move OCLC Control Number (035a) to OCLC Control Number (035z) in Network Zone records 

2. Create an Excel file that contains two columns from the DIFF tab: 001 and 035 $a. Name the file InstitutionName_035a_import. 

    * This file will be used to import/overlay OCLC Control Number (035a) in Network Zone records 
