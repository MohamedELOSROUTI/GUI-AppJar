# GUI-AppJar
The objective of this project is to build a GUI that makes it easier to send sms thanks to Sinch

## Context

For its internal activities, the AMT has to send monthly messages to its volunteers to inform them about the location of the beneficiaries and personal information (phone number, …). This way, each volunteer is able to deliver them food packages. The AMT has a set of data regarding the position of each beneficiary and its personal information saved in an excel file. The AMT desires thus to build a GUI that can help the sender to assign each volunteer to a group of beneficiaries.

## Goal
Make easy the sending of messages thanks to the GUI app offered to the used.

## Project
Build a GUI in python using AppJar library http://appjar.info/
1. The GUI should take as input:
    a. Phone number of the sender (to contact him if a problem occurs)
    b. Meeting adress (place of the food packages)
    c. Appointment date
    
2.	The GUI output:
    a.	Sending of messages containing all the beneficiaries dedicated to the given volunteer

    b.	Generate a historical file (.txt) that keeps track of all the messages that have been sent to each volunteer. Meeting adress (place of the food packages)
    
3. Procedure of the GUI:
    a. The GUI shall read the excel file containing the data of all the beneficiaries and the volunteers. 
    b.	A tableau is created containing the list of each beneficiary assigned to the volunteer
    c. The text message is automatically built according to the set of “beneficiaries” that have been crossed by the sender.
## GUI application
![center](https://i.imgur.com/UELN8JB.png)
