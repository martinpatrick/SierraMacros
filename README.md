# SierraMacros
A set of macro functions in AutoIt that automate receipt and copy cataloging functions for libraries using III Sierra on Windows 8 and above. Scripts must be compiled by the AutoIt editor to function.

The scripts here are provided without warranty, and support from Miami University Libraries is not available. These files are provided in order to form the core basis of a macro set for your needs only.

Images for the scripts are available from https://drive.google.com/folderview?id=0B9tq9izsUeT-fjlPWW1fdGl5QlBTT2l4R2twejNtdjRxOXMwcy0tWmpmZGRZdnZWMmRBTU0&usp=sharing

Folder structure:
SierraMacros/images -- contains images needed for the GUI box
SierraMacros/Scripts -- contains the scripts and the .exe compiled versions
SierraMacros/SierraGui.exe -- the GUI program that users interact with (scripts must be in SierraMacros/scripts currently as that's how they're referenced in the gui.exe)

This code was originally developed by Becky Yoose. After that it was maintained by Jason Paul Michel. The most recent contributions and updates were applied by Martin Patrick.

Originally, the receipt and copy catalog macros existed as separate file sets. Mr. Patrick combined them into a single group of files to simplify the code to allow for easier fixes and updates. Previously, any fix or update that affected both routines had to be applied to two separate files. Now, in nearly all cases, just one file needs to be fixed.

The images used by the macro GUI were created by Martin Patrick and are released into the public domain as of this notice.

One thing to note, the macros require administrator privileges to interact with OCLC on computers when the user has admin privileges. 

Files:

504 - This script inserts bibliographic references and index information into the 504 field in either Sierra or OCLC

949 - This script gets the bib record number from Sierra, and then inserts the code into a 949 field in OCLC to allow for an export and overlay for copy cataloging purposes.

BibRecord - This function runs a series of tests on the bib record to during receipt process to determine if book should go to copy cataloging team or straight to the shelf.

clearbuffer - This script ensures all variables stored in the system are cleared out. 

FindRecord - This script starts the process and searches for the book based on barcode or ISBN. Depending on which type is scanned or entered, it starts either the copy cat or receipt process.

ILOC - This script determines the item location codes for the item.

ItemCreate - This script runs a series of checks for quality control, and then proceeds to create or update an item record. This is the 'final' script that will send the item to the shelf.

multi locations item record - processes information if the bib record requires multiple location information. For instance, Miami uses two bib location codes in Sierra when a book is marked as for reference.

OCLCRec - This function searches OCLC by OCLC#, title, or ISBN to assist user in finding the record there.

OrderRecord - This script processes and enters information into the order record, and then runs the bibrecord function during the receipt process.

Special Fund List - This is a list of the endowed and donated funds that require special language to be added to the bib record. It will store these as a variable and the ItemCreate script will then enter the information appropriately to the bib record.

TSCustomFunction - This is a script that does not need to be compiled, but contains local functions developed for certain purposes and must be in the same folder as any of the other scripts when they are complied. It gets included in them at compilation.

User - A script that determines initials to enter into the bib record for tracking purposes.
