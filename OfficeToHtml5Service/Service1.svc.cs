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

namespace OfficeToHtml5Service
{
    public class Service1 : IService1
    {
        public officeConverstionDataObj getOfficeHtml5View(officeConverstionDataObj postOfficeObj)
        {
            string errorMessage = null;
            string srcPath = postOfficeObj.theFileUrl;
            string fileName = srcPath.Split('\\').Last();
            string thePhysicalFilePath = System.Web.Hosting.HostingEnvironment.MapPath("~") + @"\TempFiles\" + fileName;
            string randomFileName = DateTime.Now.ToFileTime().ToString() + "." + fileName;
            string randomFileNameWithoutExe = randomFileName.Substring(0, randomFileName.Length - (randomFileName.Split('.').Last().Length + 1));
            string fileNameWithoutExe = fileName.Substring(0, fileName.Length - (fileName.Split('.').Last().Length + 1));
            string convertedHtmlPhysicalFilePath = "", convertedHtmlFile = "";

            try
            {
                File.Copy(srcPath, thePhysicalFilePath, true);
                convertedHtmlPhysicalFilePath = thePhysicalFilePath.Replace(fileName, randomFileNameWithoutExe + ".htm");
                string applicationVirtaulRoot = System.Web.HttpRuntime.AppDomainAppVirtualPath;
                convertedHtmlFile = applicationVirtaulRoot + @"/TempFiles/" + randomFileNameWithoutExe + ".htm";

            }
            catch (Exception ex)
            {
                errorMessage += ", " + ex.Message;
            }
            errorMessage += "physical html:" + convertedHtmlPhysicalFilePath + " ,virtual html:" + convertedHtmlFile;
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
                errorMessage += "Server can not delete temperary files("+ex.Message+").";
            }
            try
            {
                //File.Copy(thePhysicalFilePath, originalFileSaveAs); 
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
                errorMessage = errorMessage + " Server can not convert the document to html("+ex.Message+").";
            }
            finally
            {
                postOfficeObj.convertedHtml5FileUrl = convertedHtmlFile;
                postOfficeObj.error = errorMessage;
            }
            return postOfficeObj;
        }
    }
}
