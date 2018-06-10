# ECMoutlook v2
- Error handling - Stops all scripts and emails IT with error


- Original email now gets copied at the end folder before the code does any other moves. This is a backup email, so nothing is lost  in servers outages. However, this creates TWO, duplicate emails at the end folder. IMO it is a good QC practice.


- New code (and icon) for entering your own QAP#. 


# ECMoutlook v1
Outlook VBA for integrating Outlook with TechnologyOne ECM. It forwards email to ECM with QAP#

The code moves emails to a working folder (zCi) under your personal email account, to accommodate for ECM not accepting emails from group email accounts.

It then adds:
a) The QAP# as a white (invisible) text and the end of the email body and
b) The original internet headers as a text at the end of the email (also white fonts)

The code forwards the formatted email to ECM and then moves it to the ending folder for QC, usually back to a group email folder for the specific user. After registering the record in ECM, those emails are expected to be deleted manually by the user.

This code expects the users to update the code with their own details (company emails etc.). The variables to be edited are clearly displayed at the start of the code.

The code is in one file: ECMOutlook v2.bas
While the ribbon/toolbar customisation with all the buttons is also included with: Outlook Customizations (olkexplorer)v2.exportedUI
