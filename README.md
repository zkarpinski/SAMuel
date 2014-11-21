SAMuel
===================
__Samuel Automates Many Unambiguous Essential Labors (still a work in progress)__

![Outlook](https://raw.github.com/zKarp/SAMuel/master/images/Outlook-preview2.JPG "Outlook tab preview with email opened.")

##Purpose
A toolset designed to automate critical tasks performed daily with my work at National Grid, during my free time at home.

##Features
 * Convert documents and images to tiffs or pdfs with click of a button (or drag 'n drop).
 * Process Emails
  * Resize attachments if larger than a letter page.
  * Compress and save the attachment as .tiff
  * Add each email, with it's corresponding account number, to an XML file to be imported.
 * Auto-generation of XML files to be imported into Kofax with the account number and batch type pre-filled.    
 * Integration with RightFax to send faxes.
 * Email documents to customers using Outlook.
 * Save user settings.

##Requirements

* .NET Framework v4.0+
* Microsoft Office 2003+
* Windows XP, Vista, 7, 8 or 8.1
* Microsoft Office Document Imaging
* RightFax 8.7+
* [Ghostscript(included)](http://www.ghostscript.com/)

##Build Requirements
* Visual Studio 2013
* [.Net Framework v4.0](http://www.microsoft.com/en-us/download/details.aspx?id=17851)
* RightFax RFCOMAPI.dll
* [Office 2003 Update: Redistributable Primary Interop](http://support.microsoft.com/kb/897646)

##Coming Soon.
 * Create database of customer emails associated with their account/customer number.
 * Email documents to customers without user interaction. (Currently creates drafts).
 * Process predefined Outlook mailboxes without user selection.
 * Auto-generate XML files without user selecting batch type or name.
 * Overhauling.

##FEEDBACK

Please submit feedback via the [issue tracker](https://github.com/zKarp/SAMuel/issues)
Copyright (c) 2014 Zachary Karpinski
