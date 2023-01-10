# PlanDay - Trent
## Sickness/Annual Leave Sync

This script will take the data that is coming from PlanDay and combine it with another Reporting Manager data that is coming from Trent.
It combine both files then it looks to extract 4 files that will be used to the Data Conversion on Trent.

* Holiday Annual Leave
* Other Absence
* Sickness Absence
* Removal of Absence

The report checks for changes done on the last report that is placed inside the /raw folder.
On this way, the changes can always be kept be up to date with the data coming in from PlanDay.

### Add New Record

If there is any changes on the data, the scrip will be picked up and create a new file to be uploaded.
### Removal Previous Record
 If there are any removal of absences that were already uploaded on a previous date, those absences will be picked up and added to the *Removal of Absence.csv*
 
 
-------------------------------
# Script Logic

vLookupManager ()
vLookupEmployeeID ()
vlookupEmployeeMail ()
vlookupManagerEMail ()

These functions are using a sort of *vlookup* method on powerhshell to find the FullName which exist on both files and return the repective value (name of function)

* AddUSID ()
This function adds a Unique String Identifier to each row, which is then use to dermine if the row needs removing or needs to be ignored

* CreateRawFiles
This creates the RAW.csv files according to their filter.(1 Annual leave, 1 Other absences, 1 Sickness)

createLoadFiles()
This is the main file which all the logic takes place, there are 3 functions inside which are nearly identical to each other, except they are looking into different files and doing different filters.

BuildSicknessDate()
This function is not currently being used, this is a separate function which looks at a constant row of dates and picks up the first and last date. For example, somebody had a holiday record like:
1,2,3,4,5,6,7 each was save into a unique row, this function would then spit out the first and last row. like so :
1
7

-------------------------------
# Running the Script
Change the *$wd* var to the location wheere the folder is extracted.

The first time it runs it will safe fail, unless there is a file on the /raw folder.
Check the files inside the folder
Files for 1st Time run
This folder can be deleted after.

-------------------------------
# Folder Logic

/raw
The files used for comparison related to the createLoadFiles() will be found inside this folder.

/.
This folder will generate "today’s file" and compare to the files that are inside the /raw

Once the run is completed it will kick up the cleanup () which will move the files into the /raw folder

/loadArea
This folder is where todays file to load into Trent will be placed

/archive
This folders are the "raw files" used to create the extract will be placed.

/log 
not currently being used

-------------------------------