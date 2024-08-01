# Hospital Preference Card Software

# Overview

The Hospital Preference Card Software is a tool designed to streamline the process of managing surgical preference cards. The software allows users to edit existing preference cards, add new services, instruments, and soft goods, and manage quantities and holds. The software is packaged as an executable file for ease of use.

# Features

<ins>Edit Existing Preference Cards:</ins> Modify existing Excel files containing preference cards.
Add New Services, Instruments, and Soft Goods: Extend the preference cards with additional items as needed.

<ins>Persistent Soft Goods Window:</ins> Open and close the soft goods window without losing inputs.
File Selection Flexibility: Select new instrument or soft goods files as long as the column headers remain consistent.

<ins>Search Functionality:</ins> Search through both the instrument and soft goods windows.
Customizable Quantities and Holds: Input any quantity and hold status for instruments and soft goods.

# Installation
<ins>Download and Extract the Zip File:</ins> Download the provided zip file, **SurveryPreferenceCardv2**, and extract its contents to your desired location.

<ins>Move the Folder:</ins> Move the entire extracted folder to your preferred location while keeping all the contents together.

<ins>Run the Executable:</ins> Navigate to the **app** folder and double-click the executable file, **run**, to run the software.

# Directory Structure

zip file name: SurgeryPrefernceCardv2 (double click to unzip)

app/ <br>
-> program/ -> [Python Program] <br>
-> data/ -> in/ -> [Instrument Input File] & [Soft Goods Input File] <br>
-> data/ -> out/ -> [All Program Outputted Doctor Preference Card Files] <br>
-> [Readme File with Instructions] <br>
-> [Executable Run File] <br>

# Usage

<ins>Editing an Existing Preference Card</ins>
Launch the Software: Double-click the executable file, 'run', inside the app folder.
Select an Input File: Choose an existing Excel file containing the preference card you wish to edit.
Modify Entries: Edit the quantities, holds, or add new services, instruments, or soft goods.
Save Changes: Save your changes to generate an updated preference card.

<ins>Adding New Entries</ins>
Select a New File: Choose a new instrument or soft goods file with consistent column headers.
Add Entries: Add new services, instruments, or soft goods as needed.
Input Details: Specify quantities and hold statuses for the new entries.
Searching and Managing Entries
Search Functionality: Use the search feature to find specific instruments or soft goods within their respective windows.
Persistent Soft Goods Window: Open the soft goods window, make selections, and close it. Reopen it to find your selections still intact.

<ins>Input and Output Files</ins>
Input Files: Ensure that input files for instruments and soft goods have consistent column headers. These files can be edited as long as the column names and file format remain unchanged.
Output Files: Generated preference cards will be saved as Excel sheets with the naming convention SurgeryName_Time_Date. Each doctor will have their own Excel file with surgeries as sheets, which will be provided to the sterile processing unit.

# Input File Headers (Case Sensitive - DO NOT CHANGE)
Instrument Container File: Service, Container Name, Reference ID <br>
Softgoods File: ITEM DESCRIPTION, VENDOR PART#

# Output File Headers (DO NOT CHANGE)
Quantity, Container Name, Iteam Description, Vendor Part #, Hold

# Notes
Ensure that the input files have the correct column headers for the program to run smoothly.
Ensure that any preference card files that need to be edited have the same column headers as generated originally.
The software generates output files in the output directory with the specified naming convention.

# Support
For any issues or support, please contact Alex West at atwest@usc.edu or alextownewest@gmail.com.
