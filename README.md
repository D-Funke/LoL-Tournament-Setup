# Tournament Sorting

## Requirements
* Python 3.9 or Later - [Download Link Included Here](https://www.python.org/downloads/)
* OpenPyXL - Version 3.0.9 or Later

The installation of python is a must in order to run the program. Once you have downloaded Python 3.9 or later, make sure to install OpenPyXL by using Python's built-in package manager pip.

Depending on your system, within a command terminal type the following command.

Note: If running into installation issues with pip, try using "sudo pip install openpyxl" for temporary root access.
### Windows
```pip install openpyxl or pip install --user openpyxl```
### Linux
```pip install openpyxl or pip install --user openpyxl```
### MacOS
```pip install openpyxl or pip install --user openpyxl```

* Note: Guaranteed to work on Windows/Linux. Unsure for MacOS.

## Usage
1. Make sure that Python is working on your system, and that you have OpenPyXL installed.

2. Download and ensure that your Excel Document follows the format as specified from the linked Google Form [here](https://docs.google.com/forms/d/1vHAIhEMqnBzDSvZyMaVmlPIyIyvCJwckU1Tpgz2PYmE/edit). Any variation in the spreadsheet formatting may result in incorrect data being produced upon sorting.

3. Place the Excel Document with your responses into the /InputForms folder. This is where the program will check for any valid documents for sorting.

```Note: It is expected that by this time you have already marked which users have "checked-in" for the tournament by specifying an X within the 'I' Column```

4. Run the "Tournament_Sort.py" program

5. Check the /SortedForms folder for your corresponding sorted excel form.


## Notes
* By default the program generates an Excel Document with 4 main spreadsheets. 
	* Raw Data - The unsorted data as provided by the form responses
	* Primary Roles - All Primary Roles sorted by Player's Rank made available for Drafting
	* Secondary Roles - All Secondary Roles sorted by Player's Rank made available for Drafting
	* User Reference - A Quick Reference made available in-case of needing to pull a player's data
