using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Text;

namespace OfficeToHtml5Service
{
    public class myCsvToHtml
    {
        private Boolean insideQuotation = false;
        private string pathFileToOpen = "";
        private List<List<string>> entries = new List<List<string>>();
        private string currentEntry = "";

        public myCsvToHtml(string _pathFileToOpen)
        {
            pathFileToOpen = _pathFileToOpen;
        }

        public String Convert()
        {
            string rst = "";
            string[] allLines = File.ReadAllLines(pathFileToOpen);
            foreach (string line in allLines)
            {
                ScanNextLine(line);
            }
            StringBuilder stb = new StringBuilder();
            /*stb.Append("<table>");
            foreach(List<string> tt in entries){
                string linestr = "";
                foreach (string valstr in tt) {
                    linestr = linestr + "<td>" + valstr + "</td>";
                }
                stb.Append("<tr>"+linestr+"</tr>");
            }
            stb.Append("</table>");*/
            foreach (List<string> tt in entries)
            {
                /*string linestr = "";
                foreach (string valstr in tt)
                {
                    linestr = linestr + "<div style='display:inline-block;float:left;'>" + valstr + "</div>";
                }*/
                stb.Append(String.Join(",",tt.ToArray()) + "\n");
                //stb.Append("<div style='clear:both;display:inline-block;'>" + linestr + "<div/>");
            }
            rst = "<div style='padding:10px;width:2000px;'><pre>"+stb.ToString()+"</pre></div>";
            return rst;
        }

        public void ScanNextLine(string line)
        {
            // At the beginning of the line
            if (!insideQuotation)
            {
                entries.Add(new List<string>());
            }

            // The characters of the line
            foreach (char c in line)
            {
                if (insideQuotation)
                {
                    if (c == '"')
                    {
                        insideQuotation = false;
                    }
                    else
                    {
                        currentEntry += c;
                    }
                }
                else if (c == ',')
                {
                    entries[entries.Count - 1].Add(currentEntry);
                    currentEntry = "";
                }
                else if (c == '"')
                {
                    insideQuotation = true;
                }
                else
                {
                    currentEntry += c;
                }
            }

            // At the end of the line
            if (!insideQuotation)
            {
                entries[entries.Count - 1].Add(currentEntry);
                currentEntry = "";
            }
            else
            {
                currentEntry += "\n";
            }
        }

    }
}