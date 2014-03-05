<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="Demo.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" lang="en">
<head runat="server">
    <title></title>
    <link href="Lib/bootstrap-3.1.1-dist/css/bootstrap.min.css" rel="stylesheet"/>
    <link href="main.css" rel="stylesheet"/>
    <script src="//code.jquery.com/jquery.js" type="text/javascript"></script>
    <script src="Lib/bootstrap-3.1.1-dist/js/bootstrap.min.js"></script>
    <script src="main.js" type="text/javascript"></script>
</head>
<body>
    <form class="form-inline" role="form" id="viewer_controller">
        <div class="form-group">
            <label for="in_filepath">Physical file path:</label>
            <input type="text" class="form-control" id="in_filepath" placeholder="c:\somefolder\somefile.someExtension" style="width:500px;"/>
        </div>
       
        <div id="btn_convert" class="btn btn-default">Convert</div>
    </form>
    <div id="myViewer"></div>
</body>
</html>
