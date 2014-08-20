## 2014-08-19

Features:

  - Kofax It: Now groups 'pages' or files with the same account number as a single document.
  - Kofax It: Added several correspondance types.
  - Add Contact: Now functions correctly using AHK script with commandline arguments.

Code Improvements:

  - Kofax It: Added dictonaries to handle each account and associated files to be uploaded together.
  - Kofax It: Now logs progress when emails canceled during processing.
  - SAM_Email Class: Added .txt as known format.

UI Improvements:

  - Kofax It: Added status text when completed.
  - General: Added welcome.
  - Convert Files: Added status text when converting.
  - Add Contact: Button disable while script is running.

## 2014-08-17

Features:

  - Outlook: Now processes emails from a selected folder rather than selected mail items.
  - Convert Files: Added PDF testing functionality using PDF995.
 
Code Improvements:

  - Email Processing: Added dispose method to SAM_Email class.
    

## 2014-08-11
Bugfixes:

  - Outlook: Fixed reject button where it still printed the images.

Features:

  - General: Added empty folders function in options.
  - General: Added view logfile button in options.
  - Outlook: Added unattended mode where emails without account numbers will be skipped automatically.
  - Convert Files: Added ability to convert image files to tiff
 
Code Improvements:

  - Email Processing: Moved all deprecated code to image processing module.
  
UI Improvements:
  - Outlook: Increased default email body font size to 10.
    

