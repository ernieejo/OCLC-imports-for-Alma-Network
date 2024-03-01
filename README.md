# OCLC-imports-for-Alma-Network
This is a series of Python scripts that can be used to sort the OCLC BIB processing report and create files for making bulk changes to the OCLC identifier in Alma records.

Prior to running these scripts, some one-time setup is needed.  

* You need to create a base folder where the scripts and the files they create will be stored.

* You may need to install the Pandas Python library

* Create an Alma Analytics with the following subjects in this order: MMS ID, OCLC Control Number (035a), OCLC Control Number (035z), Network ID. Save the Analytic as "OCLC-doublecheck"

  Tools Used:
  * OCLC WorldShare Manager
  * Alma Analytics
  * Python Scripts
  * NZ Alma (when bulk updates are needed)


<b>Step1.</b> 

* Retrieve the BIB processing report from OCLC WorldShare manager and save to your base folder. Name the file BIB_processing_report.txt 

<b>Step2.</b> 

* Run Script1 - OCLC-IZ Comparison File

This script separates BibProcessingReport.txt into two tabs in an Excel file, DIFF and SAME.​ 
* DIFF – Incoming OCLC identifier differs from what's in our system.​
* SAME – Incoming OCLC identifier is the same as what's in our system.​

1. Copy MMS ID's from the 'DIFF' tab and use them to filter the analytic 'OCLC-doublecheck'​​
2. Export the results as an Excel file and add to base file folder

<b>Step3.</b> 

* Run Script2 - OCLC-NZ Comparison File

This script combines the output from Script 1 with current information from an Alma Analytic​
* Moves any records already updated to the update tab​
* Moves any items on the do not change list to do_not_change tab
* Uses the do_not_change Excel file. Add MMS ID's for records that you want to ignore when analyzing BIB_processings_report.txt

Review the 'diff' tab​ - Often this is done manually. 
If bulk updates are needed, ensure this tab contains only rows that represent the records you want to update

<b>Step4. (Bulk updates only)</b> 

* Run Script3 - Files_for_NZ_Import
  
This script creates two files that are used to complete a bulk import of OCLC identifiers in NZ records.​
* Text file (Creat a set in Alma, then run the OCLC a_to_z normalization process on the set)

* Excel file (Use this Excel file and an NZ import profile to import the OCLC identifiers into NZ records)






