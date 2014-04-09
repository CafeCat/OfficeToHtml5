OfficeToHtml5
=============
Convert OfficeToHtml is using the office interop assemblies.
Conversion is done in the WCF service project.
convertable document types are .doc, .docx, .txt, .xls, .xlsx, .csv
<br/><br/>

How it works:<br/>
Web application take the targeted document's URL and send this to WCF application with ajax request. WCF application recieves the document's URL, copy the file to local server, convert the document to a viewable html file, and return the HTML URL to the client.
