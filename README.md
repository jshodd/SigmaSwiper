# sigmaSwiper
With many events ran by student organizations in colleges, they want to keep track of who is attending their events. This program provides a way to log all guest entries by using the student ID cards and swiping guests in. The program uses pandas to create an Excel file storing all the number of guests, the student ID of each guest, and the time of arrival. A guest list can also be loaded into the program to restrict access to an event, thus saving these organizations even more time. 

## Current Features
* Automatically check users against a guestlist (a guestlist is currently required)
* Save an excel file of all guests who attended the event 
* Extract ID number from student ID card (will require tweaking for use with other ID cards)
* Automatic emailing of list to a specified address
* Keeps a current count of the number of guests
* Displays all current guests in a scrolling list within the GUI

## Possible Future Features
* Usable without a guestlist
* Autosaving Functionality
* Integration with cloud storage (dropbox or google drive)
* Audio responses for a guest not on list
* Creating a windows installer for simple use

##Current Dependencies
* Pandas
* Openpyxl
* PyQt5
