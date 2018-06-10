#v2

- Error handling - Stops all scripts and emails IT with error


- Original email now gets copied at the end folder before the code does any other moves. This is a backup email, so nothing is lost  in servers outages. However, this creates TWO, duplicate emails at the end folder. IMO it is a good QC practice.


- New code (and icon) for entering your own QAP#. 


# ecmoutlook
Outlook VBA for integrating Outlook with ECM
ECMOutlook is the vba code and the Customisation file is for an extra toolbar with all the macros. Using this toolbar will replace any existing custimisations you might already have.

The code moves emails from shared paths to your path, as ECM will only accept emails from your own path

it then adds the QAP and the original internet headers as a twxt at the end of the email (white font)
It then emails itnto the ECM connecting email and moves it back at the shared folder but underbyour own path as a reminder that this is an email that needs to be registered. After registering it in ECM you should delete that email

The code uses a temporary placeholder folder under your account to work with the forwarding process. That folder will always be empty, but is needed for the email to originate from your account. A suggestion is to use zCi as a subfolder under your inbox

Every new set up will need to change those few variables at the top AND manually creating that zCi folder under each user

The code also assumes that the user has a shared account and a folder ubder that one for the last place those emails will stay. If not, adjust the end folder to the users send folder. Follow the style of coding as that line / Folder(folder).   
