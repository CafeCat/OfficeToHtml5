$(document).ready(function () {
    $("#btn_convert").click(function () {
        previewDocument($("#in_filepath").val());
        console.log($("#in_filepath").val());
        return false;
    });

    function previewDocument(fullUrlPath) {
        //fullUrlPath = fullUrlPath.replace(/\\/g, "*");
        $("#myViewer").empty();
        var loaderImg = $("<img></img>");
        $(loaderImg).attr("src", "images/ajax-loader-files.gif");
        $(loaderImg).attr("style", "margin-left:" + $("#myViewer").width() / 2 + "px;margin-top:" + $("#myViewer").height() / 2 + "px");
        $("#myViewer").append(loaderImg);
        if (fullUrlPath.indexOf(".doc", 0) > -1 || fullUrlPath.indexOf(".docx", 0) > -1 || fullUrlPath.indexOf(".txt", 0) > -1 || fullUrlPath.indexOf(".xls", 0) > -1 || fullUrlPath.indexOf(".xlsx", 0) > -1 || fullUrlPath.indexOf(".csv", 0) > -1) {
            var officeConverstionDataObj = { theFileUrl: fullUrlPath };
            $.ajax({
                type: "post",
                contentType: "application/json; charset=utf-8",
                url: "http://localhost/OfficeToHtml5/OfficeToHtml5Service/Service1.svc/" + "getOfficeHtml5View",
                dataType: "json",
                dataFilter: function (data) {
                    return data;
                },
                data: '{"postOfficeObj":' + JSON.stringify(officeConverstionDataObj) + '}',
                success: function (data) {
                    if (data.officeConverstionDataObj.convertedHtml5FileUrl !== null) {
                        $("#myViewer").empty();
                        if (fullUrlPath.indexOf(".csv", 0) == -1) {
                            var iframeObj = $("<iframe></iframe>");
                            iframeObj.attr("src", data.officeConverstionDataObj.convertedHtml5FileUrl);
                            iframeObj.attr("width", $("#myViewer").width());
                            iframeObj.attr("height", $("#myViewer").height() - 10);
                            iframeObj.attr("style", "border:0px");
                            $("#myViewer").append(iframeObj);
                        } else {
                            $("#myViewer").append(data.officeConverstionDataObj.fileContext);
                        }
                    } else {
                        $("#myViewer").append("<div class='alert alert-block alert-danger' style='padding:10px;margin:20px;'>" + data.officeConverstionDataObj.error + "</div>");
                    }
                },
                error: function (jqXHR, textStatus, errorThrown) {
                    $("#myViewer").append("<div class='alert alert-block alert-danger' style='padding:10px;margin:20px;'>Client request error:" + errorThrown + "</div>");
                }

            });
        } else {
            $("#myViewer").empty();
            $("#myViewer").append("<div class='alert alert-block alert-danger' style='padding:10px;margin:20px;'>This file type can not be previewed.</div>");
        }
    }
});

$(window).load(function () {
    $("#myViewer").css("height", ($(document).height() - $("#viewer_controller").height()) + 'px');
});