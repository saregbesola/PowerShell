<#
 .Synopsis
  Convert word documents to SharePoint aspx pages.
  
 .Description
  Convert word documents to SharePoint aspx pages. This function let you convert multiple word documents to web pages with a single line of command.

 .Parameter SiteUrl
  Site url.

 .Parameter TargetLibrary
  Library where documents to convert are stored.

 .Parameter Email
  Email address that can receive notification about the status of the convert operation.

 .Parameter UserName
  Username of a user who have full control access to the site and the library.
  This is used for authentication.

 .Parameter Password
  Password of the user who have acess to the site and the library.

 .Example
   # Import module.
   Import-Module ConvertWordDocument2AspxPage 

 .Example
   # Convert word to web pages.
  ConvertWordDocument2AspxPage -SiteUrl "https://domain.sharepoint.com/sites/dev" -TargetLibrary "SourceLibrary" -Email "UserName@domain.com" -UserName "UserName@domain.com" -Password "UserPassword"

 
#>
#.......................................
# Author: Aregbesola Sunday
#.......................................
  

Function ConvertWordDocument2AspxPage{
param(
    [string] $SiteUrl ,
    [string] $TargetLibrary,   
    [string] $Email,
    [string] $UserName,
    [string] $Password
  
    )

Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\DocumentFormat.OpenXml.dll";
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\OfficeDevPnP.Core.dll"
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\OpenXmlPowerTools.dll"
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\Newtonsoft.Json.dll"
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\AngleSharp.dll"


$asset=("windowsbase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35",
"DocumentFormat.OpenXml, Version=2.7.2.0, Culture=neutral, PublicKeyToken=8fb06cb64d019a17",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\Microsoft.SharePoint.Client.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\Microsoft.SharePoint.Client.Runtime.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\OfficeDevPnP.Core.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\OpenXmlPowerTools.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\System.Xml.Linq.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\System.Drawing.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\System.XML.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\Newtonsoft.Json.dll",
"C:\Program Files\WindowsPowerShell\Modules\ConvertWordDocument2AspxPage\1.0.0.0\AngleSharp.dll"

)
$src=@"

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.SharePoint.Client;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using ClientOM = Microsoft.SharePoint.Client;
using System.Security;
using OfficeDevPnP.Core;
using OpenXmlPowerTools;
using System.Drawing.Imaging;
using System.Xml.Linq;
using System.Text;
using OfficeDevPnP.Core.Pages;
using System.Threading.Tasks;
using System.Net.Mail;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

 public class WordConvertWordDocument2AspxPage
    {
      
        public static Task <string> ConvertWord2PageAsync(string SiteUrl, string SourceLibrary, string UserEmail,string UserName, string Password)
        {
          
            return Task.Run(() =>
                {
                    var docConversion=ConvertPolicyDocument(SiteUrl, SourceLibrary, UserEmail, UserName, Password);
                    return docConversion;
                });
               
        }
        private static string ConvertPolicyDocument(string SiteUrl, string SourceLibrary, string UserEmail,  string UserName, string Password)
        {
           string userName = UserName;
            string password =Password;
            try
            {
               
                SecureString securePassword = new SecureString();
                foreach (char c in password.ToCharArray())
                {
                    securePassword.AppendChar(c);
                }
                AuthenticationManager am = new AuthenticationManager();
                var htmlString = "";
                var destFileName = "";


                using (var cc = am.GetSharePointOnlineAuthenticatedContextTenant(SiteUrl, userName, securePassword))
                {       string[] filesToMerge = GetFilePath(cc, SourceLibrary);
                  for (int i = 0; i < filesToMerge.Length; i++)
                    {
                        FileInformation fileInformation =
                        ClientOM.File.OpenBinaryDirect(cc, filesToMerge[i]);            
                    
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                       
                          destFileName = filesToMerge[i].Split(new char[] { '/' }).LastOrDefault().Replace(".docx", ".aspx");
                            var destLib = filesToMerge[i].Split(new char[] { '/' });
                            var pageTitle = destFileName.Replace(".aspx", "");
                            var imageDirectoryName = filesToMerge[i].Substring(0, filesToMerge[i].Length - filesToMerge[i].Split(new char[] { '/' }).LastOrDefault().Length);
                            var imageDirectoryRelativeName = destFileName.Replace(".aspx", "").Trim() + "_Images";
                        List list = cc.Web.Lists.GetByTitle(SourceLibrary);
                        FolderCollection folders = list.RootFolder.Folders;
                        cc.Load(folders);
                        cc.ExecuteQuery();
                        Folder ImagefolderExist = folders.Where(e => e.Name == imageDirectoryRelativeName).FirstOrDefault();
                        if (ImagefolderExist == null)
                        {
                            ListItemCreationInformation info = new ListItemCreationInformation();
                            info.UnderlyingObjectType = FileSystemObjectType.Folder;
                            info.LeafName = imageDirectoryRelativeName.Trim();//Trim for spaces.Just extra check
                            ListItem newItem = list.AddItem(info);
                            newItem["Title"] = imageDirectoryRelativeName;
                            newItem.Update();
                        }
                        CopyStream(fileInformation.Stream, memoryStream);
                        using (WordprocessingDocument doc =
                                                WordprocessingDocument.Open(memoryStream, true))
                        {
                            byte[] streamArr = memoryStream.ToArray();
                            bool fixedSize = streamArr.IsFixedSize;
                            DocumentFormat.OpenXml.Packaging.HeaderPart headerText = doc.MainDocumentPart.HeaderParts.FirstOrDefault();
                            //DocumentFormat.OpenXml.Packaging.FooterPart footerText = doc.MainDocumentPart.FooterParts.FirstOrDefault();

                            int imageCounter = 0;

                            if (headerText != null)
                            {
                                pageTitle = headerText.Header.InnerText;
                            }
                            else
                            {
                                var part = doc.CoreFilePropertiesPart;
                                if (part != null)
                                {
                                    pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? pageTitle.Trim();
                                }
                                else
                                    pageTitle = "No Title";
                            }
                            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                            {
                                PageTitle = pageTitle,
                                FabricateCssClasses = true,
                                CssClassPrefix = "pt-",
                                RestrictToSupportedLanguages = false,
                                RestrictToSupportedNumberingFormats = false,
                                ImageHandler = imageInfo =>
                                {
                                    ++imageCounter;

                                    string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                                    ImageFormat imageFormat = null;
                                    if (extension == "png")
                                        imageFormat = ImageFormat.Png;
                                    else if (extension == "gif")
                                        imageFormat = ImageFormat.Gif;
                                    else if (extension == "bmp")
                                        imageFormat = ImageFormat.Bmp;
                                    else if (extension == "jpeg")
                                        imageFormat = ImageFormat.Jpeg;
                                    else if (extension == "tiff")
                                    {
                                        // Convert tiff to gif.
                                        extension = "gif";
                                        imageFormat = ImageFormat.Gif;
                                    }
                                    else if (extension == "x-wmf")
                                    {
                                        extension = "wmf";
                                        imageFormat = ImageFormat.Wmf;
                                    }

                                    // If the image format isn't one that we expect, ignore it,
                                    // and don't return markup for the link.
                                    if (imageFormat == null)
                                      return null;                          

                                    try
                                    {
                                        using (MemoryStream ms = new MemoryStream())
                                        {
                                            imageInfo.Bitmap.Save(ms, imageFormat);
                                            ms.Seek(0, SeekOrigin.Begin);
                                            if (cc.HasPendingRequest)
                                                cc.ExecuteQuery();
                                            try
                                            {
                                                ClientOM.File.SaveBinaryDirect(cc, imageDirectoryName + imageDirectoryRelativeName + "/image" + imageCounter + "." + imageFormat, ms, true);
                                            }
                                            catch
                                            {
                                                try
                                                {
                                                    ClientOM.File.SaveBinaryDirect(cc, imageDirectoryName + imageDirectoryRelativeName + "/image" + imageCounter + "." + imageFormat, ms, true);
                                                }
                                                catch
                                                {

                                                }
                                            }

                                        }
                                    }
                                    catch (System.Runtime.InteropServices.ExternalException)
                                    {
                                        return null;
                                    }
                                    XElement img = new XElement(Xhtml.img,
                                        imageInfo.ImgStyleAttribute,
                                        imageInfo.AltText != null ?
                                            new XAttribute(NoNamespace.alt, imageInfo.AltText) : null,
                                        new XAttribute(NoNamespace.src, imageDirectoryName + imageDirectoryRelativeName + "/image" + imageCounter + "." + imageFormat)
                                        );
                                    return img;
                                }
                            };
                            XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(doc, settings);

                            // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                            // we are using HTML5.
                            var html = new XDocument(
                                new XDocumentType("html", null, null, null),
                                htmlElement);

                            htmlString = html.ToString(SaveOptions.DisableFormatting);

                        }

                        var page = cc.Web.AddClientSidePage(destFileName, true);
                     
                        ClientSideText txt2 = new ClientSideText() { Text = htmlString };
                        
                        page.AddControl(txt2, -1);
                        page.PageTitle = pageTitle;
                        page.Save();
                        page.Publish();

                    }
                    SendEmails(UserEmail, destFileName+" Page successfully created.",destFileName+" page has been successfully created. You can find it in the Site Pages library at the destination site." , userName, password);
                  }
                }
                 
                return "Document conversion operation successfully completed.";
            }
            catch(Exception ex)
            {

                SendEmails(UserEmail, "Policy page failed to create", ex.ToString(), userName, password);
                return ex.ToString();
            }
        }
        private static string[] GetFilePath(ClientContext cc, string sourceLibrary)
        {
            List<string> buildFilePath = new List<string>();
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View Scope='Recursive'>
                <Query>
                  <Where>
                    <Eq>
                      <FieldRef Name='File_x0020_Type'/><Value Type='Text'>docx</Value>
                    </Eq>
                  </Where> 
                 </Query>
              </View>";
            var web = cc.Web;
            List list = cc.Web.Lists.GetByTitle(sourceLibrary);
            ListItemCollection files = list.GetItems(query);

            cc.Load(list);
            cc.Load(files, file => file.Include(filePop => filePop.File));
            cc.ExecuteQuery();

            foreach (var file in files)
            {

                var filePath = file.File.ServerRelativeUrl;

                buildFilePath.Add(filePath);

            }

            string[] filePaths = buildFilePath.ToArray();

            string[] documentsToMerge = filePaths;

            return documentsToMerge;
        }
        protected static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;
            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }
        private static string DecodeFrom64(string encodedData)
        {
            System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
            System.Text.Decoder utf8Decode = encoder.GetDecoder();
            byte[] todecode_byte = Convert.FromBase64String(encodedData);
            int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
            char[] decoded_char = new char[charCount];
            utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
            string result = new String(decoded_char);
            return result;
        }

        private static void SendEmails(string _To, string _subject, string _textBody,string userName, string password)
        {
            // Create and build a new MailMessage object
            MailMessage message = new MailMessage();
            message.IsBodyHtml = true;
            message.From = new MailAddress(userName, "SP Admin");
            message.To.Add(new MailAddress(_To));
            message.Subject = _subject;
            message.Body = _textBody;
            {
                using (SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.office365.com",
                    Port = 587,
                    Credentials = new System.Net.NetworkCredential(userName, password),
                    EnableSsl = true
                }
                )
                {
                    try { smtp.Send(message); }
                    catch (Exception excp)
                    {
                        Console.Write(excp.Message);
                     
                    }
                }
            }

        }
    }
"@



 Add-Type -ReferencedAssemblies $asset -TypeDefinition $src
 
 [WordConvertWordDocument2AspxPage]::ConvertWord2PageAsync($SiteUrl,$TargetLibrary,$Email,$UserName,$Password)

 }

 Export-ModuleMember -Function ConvertWordDocument2AspxPage
