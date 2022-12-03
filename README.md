To get started, download or clone the project file. If you download it as a zip file, unzip it and copy the folder named ConvertWordDocumentToModernPage along with its content to Modules folder inside WindowsPowerShell directory in Program Files i.e. C:\Program Files\WindowsPowerShell\Modules.  Then follow the steps below to import the module and run the command.

1. Open PowerShell ISE or PowerShell command window 
2. Copy and paste the command below:

   Import-Module ConvertWordDocumentToModernPage 
   
   ConvertWordDocumentToModernPage -SiteUrl " " -TargetLibrary " " -UserName " " -Password " " -FileName " "

Supply value for each of the parameter as explained below:

SiteUrl - Url of the site where the word document are located

TargetLibrary - Library where the documents are stored

UserName - Username or email address of a user who have full control to the site and the library. This is used for authentication

Password - Password of the account used for authentication

Note: if you don't have access to Program Files. Copy the folder i.e. ConvertWordDocumentToModernPage  to any location of your choice. Then navigate to inside of the folder, and open the ConvertWordDocumentToModernPage.psm1 file. Change path of the dlls files to the location where you have copied the folder. 

Open PowerShell command and change the directory to where you have the ConvertWordDocument2AspxPage  and import the module using the command below:
Import-Module (Resolve-Path('ConvertWordDocumentToModernPage'))

Note: If you receive Unable to load one or more of the requested types error, please ignore.
