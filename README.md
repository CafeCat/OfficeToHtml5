OfficeToHtml5
=============
Convert OfficeToHtml is using the office interop assemblies.
Samples are presented in the WCF service.
Currently it's set up for debugger only, so the end point binding is not webHttpBinding.

To test:
1. use debugger.
2. enter a file physical path value in the property theFileURL(I know, it's not a url, it's a long story.). 
3. Invoke
4. get the viewable url path from the property convertedHtml5Url
