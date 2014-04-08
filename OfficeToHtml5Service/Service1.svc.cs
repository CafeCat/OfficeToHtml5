using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net;

namespace OfficeToHtml5Service
{
    public class Service1 : IService1
    {
        public officeConverstionDataObj getOfficeHtml5View(officeConverstionDataObj postOfficeObj)
        {
            string errorMessage = null;
            string fileName = postOfficeObj.theFileUrl.Split('/').Last();
            string thePhysicalFilePath = System.Web.Hosting.HostingEnvironment.MapPath("~") + @"\TempFiles\" + fileName;
            string fileNameWithoutExe = fileName.Substring(0, fileName.Length - (fileName.Split('.').Last().Length + 1));
            string randomFileNameWithoutExe = DateTime.Now.ToFileTime().ToString() + "." + fileNameWithoutExe;
            string convertedHtmlPhysicalFilePath = thePhysicalFilePath.Replace(fileName, randomFileNameWithoutExe + ".htm");
            string convertedHtmlUrl = System.Web.HttpRuntime.AppDomainAppVirtualPath + @"/TempFiles/" + randomFileNameWithoutExe + ".htm";
            //postOfficeObj.debugMsg = "get directory name:" + Path.GetDirectoryName(convertedHtmlPhysicalFilePath) + " nameEnding:" + convertedHtmlPhysicalFilePath.Replace(convertedHtmlPhysicalFilePath.Split('.')[0] + ".", "");

            //Delete converted temp html files
            try
            {
                string originalFileNameEnding = convertedHtmlPhysicalFilePath.Replace(convertedHtmlPhysicalFilePath.Split('.')[0] + ".", "");
                DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(convertedHtmlPhysicalFilePath));
                foreach (FileInfo theFile in dir.GetFiles())
                {
                    if (theFile.Name.Contains(originalFileNameEnding))
                    {
                        theFile.Delete();
                    }
                }
                foreach (DirectoryInfo theDirectory in dir.GetDirectories())
                {
                    if (theDirectory.Name.Contains(fileNameWithoutExe + "_files"))
                    {
                        theDirectory.Delete(true);
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = "Server can not delete temperary files.";
            }

            //Copy file from remoate location to FeedReports's temp folder
            using (WebClient wc = new WebClient())
            {
                wc.UseDefaultCredentials = true;
                wc.DownloadFile(postOfficeObj.theFileUrl, thePhysicalFilePath);
            }


            //Convert documents
            if (File.Exists(thePhysicalFilePath))
            {
                try
                {
                    string ext = fileName.Split('.').Last();
                    if (ext == "doc" || ext == "docx" || ext == "txt")
                    {
                        myDocToHtml docToHtml = new myDocToHtml(thePhysicalFilePath, convertedHtmlPhysicalFilePath);
                        docToHtml.convert();
                    }
                    else if (ext == "xls" || ext == "xlsx")
                    {
                        object XlsFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml;
                        myXlsToHtml myXlsHtml = new myXlsToHtml(thePhysicalFilePath, convertedHtmlPhysicalFilePath);
                        myXlsHtml.convert();
                    }
                    else if (ext == "csv")
                    {
                        myCsvToHtml csvToHtml = new myCsvToHtml(thePhysicalFilePath);
                        postOfficeObj.fileContext = csvToHtml.Convert();
                    }
                }
                catch (Exception ex)
                {
                    errorMessage = errorMessage + " Server can not convert the document to html.";
                }
            }
            else
            {
                errorMessage += "Copy files from remote server failed.";
            }
            postOfficeObj.convertedHtml5FileUrl = convertedHtmlUrl;
            postOfficeObj.error = errorMessage;
            return postOfficeObj;
        }
    }
}
