# ArchivesSpaceUBLAPI
Scripts to retrieve data from the ArchivesSpace API 

# Requirements
- Python verson 3.1 or higher.
- Connected to the UBL network through either LAN connection or NUWD-laptop WiFi.
- ArchivesSpace account with Admin rights.

# singleItemGet
Retrieves a single Archival Object. Run with --obj=<uri of Archival Object> to retrieve that object as jSon. 

# getResObjData
Retrieves all Archival Objects part of a Resource.
Run with --ubl=<ubl number of EAD> to specify the Resource.
Writes uri, ref_ID, level, title, component_ID, notes, and subnotes to an Excel file.
Order in Excel file is seemingly random and needs sorting in Excel itself.
