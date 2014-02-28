using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace OfficeToHtml5Service
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IService1
    {
        public officeConverstionDataObj getOfficeHtml5View(officeConverstionDataObj postOfficeObj)
        {
            string testPath = postOfficeObj.theFileUrl;
            string fileName = postOfficeObj.theFileUrl.Split('/').Last();
            string randomFileName = DateTime.Now.ToFileTime().ToString() + "." + postOfficeObj.theFileUrl.Split('/').Last();
            string extratUsefulPath = testPath.Replace(testPath.Split('/')[0] + "/" + testPath.Split('/')[1] + "/" + testPath.Split('/')[2] + "/", "/");
            string fileNameWithoutExe = fileName.Substring(0, fileName.Length - (fileName.Split('.').Last().Length + 1));
            string randomFileNameWithoutExe = randomFileName.Substring(0, randomFileName.Length - (randomFileName.Split('.').Last().Length + 1));
            string thePhysicalFilePath = System.Web.Hosting.HostingEnvironment.MapPath(extratUsefulPath);
            string convertedHtmlPhysicalFilePath = thePhysicalFilePath.Replace(fileName, randomFileNameWithoutExe + ".html");
            string convertedHtmlFile = testPath.Replace(fileName, randomFileNameWithoutExe + ".html");
            postOfficeObj.debugMsg = "get directory name:" + Path.GetDirectoryName(convertedHtmlPhysicalFilePath) + " nameEnding:" + convertedHtmlPhysicalFilePath.Replace(convertedHtmlPhysicalFilePath.Split('.')[0] + ".", "");
            string errorMessage = null;
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
                errorMessage = errorMessage + " Server can not convert the document to html.";
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
