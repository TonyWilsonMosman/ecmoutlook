# ECMoutlook v2.85 

Minimising even more set up. 

Needs changing email addresses of the main account and of the IT support account 

And changing the folder path of when the emails should end up for QC. The original email ends up there. The copy send to ECM ends up in the users deleted folder 

  

-wish list a) adding all variable upfront for easy installation b) internet headers are not 100% complete, missing some IDs 

  

# ECMoutlook v2.85 
2.5 is a test variation of v2. It simplifies initial set up and needs testing before deployment. It has issues recognising the first user variable, users email account. It works if you manually replace it with the user email. 

  

  

# ECMoutlook v2 

- Error handling - Stops all scripts and emails IT with error 

  

- Original email now gets copied at the end folder before the code does any other moves. This is a backup email, so nothing is lost in servers outages. However, this creates TWO, duplicate emails at the end folder. IMO it is a good QC practice. 

  

- New code (and icon) for entering your own QAP#.  

  

***Note: There is an awesome alternative of an add-on by Tony Federer (https://github.com/tonyfederer), which needs no customisation per user and it possibly the best option for organisation-wide deployment. However, this code is good for teams that register a lot of records and would like to use QAP# for quick and direct registration, and where IT is around to assist with deployment at each user. 

  

# ECMoutlook v1 

Outlook VBA for integrating Outlook with TechnologyOne ECM. It forwards email to ECM with QAP# 

  

The code moves emails to a working folder (zCi) under your personal email account, to accommodate for ECM not accepting emails from group email accounts. 

  

It then adds: 

- The QAP# as a white (invisible) text and the end of the email body and 

- The original internet headers as a text at the end of the email (also white fonts) 

  

The code forwards the formatted email to ECM and then moves it to the ending folder for QC, usually back to a group email folder for the specific user. After registering the record in ECM, those emails are expected to be deleted manually by the user. 

  

This code expects the users to update the code with their own details (company emails etc.). The variables to be edited are clearly displayed at the start of the code. 

  

The code is in one file: "ECMOutlook v2.bas", while the ribbon/toolbar customisation with all the buttons is also included with "Outlook Customizations (olkexplorer)v2.exportedUI" 

 
