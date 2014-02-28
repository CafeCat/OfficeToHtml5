using System;
using System.Collections;
using System.Collections.Generic;
using System.Web;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Text;

namespace OfficeToHtml5Service
{
    public class myDocToHtml
    {
        private string path = "";
        private object outputFileName = null;
        public myDocToHtml(string _path, object _outputFileName) {
            path = _path;
            outputFileName = _outputFileName;
        }
        public void convert() {
            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            try
            {
                word.Visible = false;
                word.ScreenUpdating = false;
                try
                {
                    Documents oDocTmp = word.Documents;
                    word.Visible = false;

                    Document doc = oDocTmp.Open((Object)path, ref oMissing,
                                                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    doc.Activate();

                    //object outputFileName = outputFieName;
                    //outputFileName = path.Replace(".doc", ".html");


                    object fileFormat = WdSaveFormat.wdFormatHTML;

                    doc.SaveAs(ref outputFileName,
                                ref fileFormat, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                }
                finally
                {
                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    // ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                    // doc = null;
                }
            }
            finally
            {
                ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                word = null;
            }
        }
    }
}