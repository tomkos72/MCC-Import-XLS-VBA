# MCC-Import-XLS-VBA
Excel with VBA to import downloadable MCC's and create a summary Excel

In the context of the internal control system, individual control catalogues are provided annualy for different processes.
To obtain an overall global overview of all risks and control objectives in one Excel file, this excel and the embedded VBA code has been created to support the process.
The global Excel file contains a summary page and a Pivot table page, that allows easy filtering/searches.

One or more MCC Excel files can be selected, to be inserted/included in that global Excel. The Summary- and Pivot-Page will be updated accordingly.

Preconditions:
- all MCC's are downloaded to a local/network directory
- the MCC files follow the naming guidelines, containing leading: "MCC_" and the next 3 letters are the process abreviation in upper case (e.g. ACC for Accounting)
- each MCC file contains a sheet "MCC", that contains the data
- the MCC sheet contains in line 1 and 2 the title information
- the MCC sheet contains from line 3 onward the data
- the MCC sheet contains data beween row A and row W, where the "Master Process Shortname" is in column "B" in upper case. The same as used in the filename.

In case the MCC format or the MCC naming conventions change, the VBA code must be updated accordingly.
