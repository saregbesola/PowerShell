To get started, download or clone the project file. If you download it as a zip file, unzip it and copy the folder named ConvertWordDocument2AspxPage along with its content to Modules folder inside WindowsPowerShell directory in Program Files i.e. C:\Program Files\WindowsPowerShell\Modules.  Then follow the steps below to import the module and run the command.

1. Open PowerShell ISE or PowerShell command window 
2. Copy and paste the command below:

   Import-Module ConvertWordDocument2AspxPage 
   ConvertWordDocument2AspxPage -SiteUrl " " -TargetLibrary " " -Email " " -UserName " " -Password " "

Supply value for each of the parameter as explained below:

SiteUrl - Url of the site where the word document are located

TargetLibrary - Library where the documents are stored

Email - Email address that will receive notification when the page conversion finishes or error occurs

UserName - Username or email address of a user who have full control to the site and the library. This is used for authentication

Password - Password of the account used for authentication

Note: if you don't have access to Program Files. Copy the folder i.e. ConvertWordDocument2AspxPage  to any location of your choice. Then navigate to inside of the folder, and open the ConvertWordDocument2AspxPage.psm1 file. Change path of the dlls files to the location where you have copied the folder. 

Open PowerShell command and change the directory to where you have the ConvertWordDocument2AspxPage  and import the module using the command below:
Import-Module (Resolve-Path('ConvertWordDocument2AspxPage'))

