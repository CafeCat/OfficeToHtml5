using System;
using System.Collections;
using System.Collections.Generic;
using System.Web;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using System.Text;

namespace OfficeToHtml5Service
{
    public class myXlsToHtml
    {
        private string pathFileToOpen = "";
        private string pathFileToSave = "";
        public myXlsToHtml(string _pathFileToOpen, string _pathFileToSave) {
            pathFileToOpen = _pathFileToOpen;
            pathFileToSave = _pathFileToSave;
        }
        public Boolean convert() 
        {
            Boolean rst = false;
            Application excel = new Application();
            
            try
            {
                excel.Workbooks.Open(Filename: pathFileToOpen);
                excel.Visible = false;
                if (excel.Workbooks.Count > 0)
                {
                    IEnumerator wsEnumerator = excel.ActiveWorkbook.Worksheets.GetEnumerator();

                    object format = Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml;
                    int i = 1;
                    while (wsEnumerator.MoveNext())
                    {
                        Microsoft.Office.Interop.Excel.Worksheet wsCurrent = (Microsoft.Office.Interop.Excel.Worksheet)wsEnumerator.Current;
                        String outputFile = "excelFile" + "." + i.ToString() + ".html";
                        wsCurrent.SaveAs(Filename: pathFileToSave, FileFormat: format);
                        ++i;
                        break;
                    }
                    excel.Workbooks.Close();
                }
            }
            catch (Exception ex)
            {
                rst = false;
            }
            finally
            {
                excel.Application.Quit();
                rst = true;
            }
            return rst;
        }
        public StringBuilder readConvertedFile()
        {
            StringBuilder ConvertedResult = new StringBuilder();
            string fileName = Path.GetFileName(pathFileToSave);
            string directory = Path.GetDirectoryName(pathFileToSave) + "\\" + fileName.Split('.')[0] + "_files";
            string[] files = Directory.GetFiles(directory);
            for (int i = 0; i < files.Length - 1; i++)
            {
                if (Path.GetExtension(files[i]) == ".html")
                {
                    ConvertedResult.Append(File.ReadAllText(files[i]).Replace("<![endif]>", ""));
                }
            }
            return ConvertedResult;
        }

    }
}