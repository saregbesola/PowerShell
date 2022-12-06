# Description
  Convert word documents to SharePoint modern pages. This function let you convert a single or multiple word documents to modern pages with a single line of command.
  
# Import module
To get started, download or clone the repository. If you download it as a zip file, unzip it and copy the folder named ConvertWordDocumentToModernPage along with its content to Modules folder inside WindowsPowerShell directory in Program Files i.e. C:\Program Files\WindowsPowerShell\Modules.  Then follow the steps below to import the module and run the command.

1. Open PowerShell ISE or PowerShell command window 
2. Copy and paste the commands below:

   Import-Module ConvertWordDocumentToModernPage 
   
   ConvertWordDocumentToModernPage -SiteUrl " " -TargetLibrary " " -UserName " " -Password " " -FileName " "

Supply value for each of the parameters as explained below:

SiteUrl - Url of the site where the word document are located

TargetLibrary - Library where the documents are stored

UserName - Username or email address of a user who has full control on the site and the library. This is used for authentication and accessing the files.

Password - Password of the account used for authentication.

FieName - Optional: specify a filename for a word document to convert. If you do not use FileName parameter all word documents in the library will be converted. 

# Example
  To convert a single word document: ConvertWordDocumentToModernPage -SiteUrl "https://domain.sharepoint.com/sites/dev" -TargetLibrary "SourceLibrary" -UserName "UserName@domain.com" -Password "UserPassword" -FileName "ConvertWord.docx"
  
  To convert all word documents in a library:  ConvertWordDocumentToModernPage -SiteUrl "https://domain.sharepoint.com/sites/dev" -TargetLibrary "SourceLibrary" -UserName "UserName@domain.com" -Password "UserPassword"
  
Note: if you don't have access to Program Files. Copy the folder i.e. ConvertWordDocumentToModernPage  to any location of your choice. Then navigate to inside of the folder, and open the ConvertWordDocumentToModernPage.psm1 file. Change path of the dlls files to the location where you have copied the folder. 

Open PowerShell command and change the directory to where you have the ConvertWordDocument2AspxPage  and import the module using the command below:
Import-Module (Resolve-Path('ConvertWordDocumentToModernPage'))

# Possible error you might receive
1. If you receive Unable to load one or more of the requested types error, please ignore.

2. ConvertWordDocumentToModernPage.psm1 is not digitally signed. You cannot run this script on the current system. Run the two commands below (note if you put the ConertWordDocumentToModernPage in a different directory, change the path to that directory)
  
  . Unblock-File -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocumentToModernPage\ConvertWordDocumentToModernPage.psm1"

  . dir "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocumentToModernPage\1.0.0.0" | Unblock-File

