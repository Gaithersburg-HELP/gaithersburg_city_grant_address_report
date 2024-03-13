# Gaithersburg City Grant and Montgomery County Totals Report
## About
This Excel macro workbook takes in visit data records exported as a CSV from food bank manager software such as Soxbox and produces the quarterly address listings report for Gaithersburg City as well as the monthly Montgomery County visit totals report.

### Overview:
1) Imports visit data records
2) Groups visit data records together by address
3) Attempts to validate addresses against the Gaithersburg Address database
4) Can additionally validate addresses against the Google Address Validation API
5) Can additionally accept user input to fix invalid addresses
6) Produces visit totals as well as the Gaithersburg City address listings report

# Using the XLSM file
## Downloading the XLSM
1) Ensure you have the latest version. Download the [latest release of the XLSM file](https://github.com/jimmyli97/gaithersburg_city_grant_address_report/releases). Click on the "Assets" title and then click on the XLSM to download. 
![Release assets download page](readme/1download.png)
    * If you have data in an older release version, in the new file on the "Interface" sheet, click "Import Data" and select the older file. All data will be copied over to the new version.
2) The same file can be used from year to year and from quarter to quarter. The XLSM file will remember previously validated and user edited addresses.
3) If this is the first quarter, name the file with the current fiscal year, "e.g. City Grant Address Listings Report v3.0 FY24.xlsm"
    * If you have a file from the last fiscal year, make a copy of it for this fiscal year and rename it. Then on the "Addresses" sheet, click "Delete All Visit Data" to delete all visit data but keep address data.
## Importing data
1) Log into your food bank manager and export data as a CSV. The visit data does not need to be quarterly, the XLSM file will automatically sort by quarter. The visit data can also be imported at any time, you don't have to do it all at once at the end of the quarter.
    1) For Gaithersburg HELP Soxbox, log in [here](https://ghp.soxbox.co/login). Go to Visit History Export:
       ![Soxbox Visit History Export](readme/2.1soxbox_visithistoryexport.png)
    2) Select the dates you wish to export. For instance, to run the county monthly totals for this month, export dates for this month only. Select the preset "city and county grant address v3". Click "Export" and save the CSV file
       ![city and county grant address v3 preset and date selection](readme/2.2soxbox_preset.png)
2) Open the XLSM file. If you see the Protected View warning message, click the "Enable Editing" button. If you see that macros from the internet are disabled, close the workbook, right click on the workbook in File Explorer, go to Properties, at the bottom of the General tab check the Unlock checkbox, and open the workbook again (see [this link](https://learn.microsoft.com/en-us/deployoffice/security/internet-macros-blocked)).
6) Google Address Validation requires a [Google Address Validation key](https://developers.google.com/maps/documentation/address-validation/get-api-key). This file expects a file named "apikeys.csv" formatted as "address_key,apikey", placed in the same directory as the XLSM.

