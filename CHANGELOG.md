## 2014-08-29

Features:

  - Outlook: Removed picture box. Replaced by listview.
   -Outlook: Added listview of current email attachments. Clicking the entry, opens the file with the default program.
   -RightFaxIt: Added polling feature for one folder. (Testing)

Code Improvements:

  - Convert Files: Release sets MODI as active printer, Debug does not. (Since I don't have MODI on my machine)
  - General: Added new SAM_File class to be used for easier file attribute references.
  - Outlook: SAM_Email's use the new SAM_File class
  - Outlook: Groups all attachments from an email as one audit. Previously SAMuel audited each individual attachment within an email.
  - RightFaxIt: Skips generic phone number files like 1-999-999-9999.

UI Improvements:

  - Outlook: Removed a hidden label.
  - Outlook: Removed picturebox.

## 2014-08-29

Code Improvements:

  - Convert Files: Removed duplicate code.
  - Outlook: Imagebox on click, opens non-images now. Preview still in the works.

## 2014-08-26

Features:

  - Outlook: Now handles pdf files using [Ghostscript](http://ghostscript.com/).

Code Improvements:

  - Convert Files: Word documents are opened as read-only.
  - Outlook: Images now use same function as the convert files tab.

UI Improvements:

  - RightFaxIt: Removed the completed faxing messagebox.

## 2014-08-24

Features:

  - Add Contact: Add contacts to multiple account by parsing file names with drag 'n drop. (Still in progress.) 

Code Improvements:

  - RightFaxIt: Checks if file to be move exists in destination. If it does, delete it then move.
  - Email Processing: Cleaned Regex code.
  - General: Blocked of code depending on the tab it belonged to.

UI Improvements:

  - General: Converting with printer no longer shows printing x of y dialog.
  - RightFaxIt: Removed access to options until it's useful.

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
    

