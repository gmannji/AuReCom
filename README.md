# Automated Risk Register Review and Compare (AuReCom)

This project is created as a tool for my monthly project risk reviews. Currently attached to a construction program where I need to manage 10 risk registers, in other words, 10 projects. It is easier to create a script that will do it for you although learning python (even coding) as you code is quite a challenge.

As this is my first program, there are many inconsistencies as i was exploring or experimenting new ways to execute what i wanted.

The first part of AuReCom is the risk register review while the second part will compare the risk register to previous month and identify what has changed and what has been accomplished.

Initially, the idea came after i learned using openpyxl and straight started this project. Then stumbled upon using pandas for the second part. The project started with manually assign specific cells in excel for checking and review then transition to automatically select which cell to look at, identify the header row and etc.

The script will initially request for an excel file downloaded from our Risk Management System in a specific format as first input. Then it'll start to flag the risk register for errors and cells or items that require attention (risks triggering soon, risks expired and expiring soon, mitigation actions due and overdue, etc). It'll then write a status update in a new cell next to the end column of first row of the current risk register by giving the number of risks, risk status, and actions status. 

New excel file created and saved.

Then it'll request for another excel file for comparison, this is usually previous month's risk register that i am reviewing. The reason for comparing to last month's risk register is that i can track what have the projects accomplished in their risk management activities. Changes will be captured in a new sheet and distinguished with "old --<<>>-- new" format. At the same time changes that i care about (i.e changes to mitigation actions) are interpreted and added to a new column in this new sheet.

During this stage, the script will also identify which risks are new and which are closed or retired at the time of review in new sheets.

Finally, all of the changes that i care about will be shown in the same cell as the status update.

Then saved on the new excel file created.

Additionally the script is converted to .exe so my other teammates can make a good use out of it.
