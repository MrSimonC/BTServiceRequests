# BT Service Requests
For NHS trusts in the national program, it is common to log service requests to BT to solve 3rd line support issues with Cerner Millennium. Where I work this also involves creating a call on our helpdesk software SDPlus and adding an entry to an excel sheet.

This program allows you to quickly log common service requests, by automating the browser actions of logging a normal service request.
This python program will open a browser, log you in to the BT service request page, browse to the appropriate service request number, filling all the details then will press order now, pausing at the submit stage.
Once you're happy with the order you can press submit, then read off the RITM number ready to put into service desk plus via the SDPlus menu.