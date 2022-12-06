# Description
  This PowerShell module creates SharePoint modern pages from word documents (docx). This function let you convert a single or multiple word documents to modern pages with a single line of command. This tool is for you if you have wanted an easy way to create table of content (TOC) in modern pages without installing a 3rd party web part, include Back to top functionality in your page or convert your existing documents to web pages e.g. policy documents. Just add TOC or back to top tag in your word document and you are set to run the command. 
  
  A sample page created with table of content along with anchor link to each heading/section.
  ![image](https://user-images.githubusercontent.com/35005799/205815181-1fbbd405-ec81-4ed4-aec9-5350e3d1da06.png)

  

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

FieName - Optional: specify a filename for a word document to convert. If you do not use the FileName parameter, modern pages will be created for all word documents (docx) in the library. 

# Example
  To convert a single word document: ConvertWordDocumentToModernPage -SiteUrl "https://domain.sharepoint.com/sites/dev" -TargetLibrary "SourceLibrary" -UserName "UserName@domain.com" -Password "UserPassword" -FileName "ConvertWord.docx"
  
  To convert all word documents in a library:  ConvertWordDocumentToModernPage -SiteUrl "https://domain.sharepoint.com/sites/dev" -TargetLibrary "SourceLibrary" -UserName "UserName@domain.com" -Password "UserPassword"
  
Note: if you don't have access to Program Files. Copy the folder i.e. ConvertWordDocumentToModernPage  to any location of your choice. Then navigate to inside of the folder, and open the ConvertWordDocumentToModernPage.psm1 file. Change path of the dlls files to the location where you have copied the folder. 

Open PowerShell command and change the directory to where you have the ConvertWordDocument2AspxPage  and import the module using the command below:
Import-Module (Resolve-Path('ConvertWordDocumentToModernPage'))

# Possible errors you might receive
1. If you receive "Unable to load one or more of the requested types" error, please ignore.

2. You might also receive "ConvertWordDocumentToModernPage.psm1 is not digitally signed. You cannot run this script on the current system." This is beacuse you have RemoteSigned policy enabled that prevents you from running scripts that are downloaded from the internet unless they are digitally signed. Start PowerShell as an Administrator and run the two commands below to unblock the downloads (note if you put the ConvertWordDocumentToModernPage in a different directory, change the path to that directory):
  
   Unblock-File -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocumentToModernPage\ConvertWordDocumentToModernPage.psm1"

   dir "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocumentToModernPage\1.0.0.0" | Unblock-File
3. Close and re-open the PowerShell console then try again

# Caveat
You will lose the formatting if you edit the page directly. if you need to update the page, make the changes in the word document and run the command again. The best way is to automate publishing the changes using the command. 
