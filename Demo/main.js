$(document).ready(function () {
    $("#btn_convert").click(function () {
        previewDocument($("#in_filepath").val());
        console.log($("#in_filepath").val());
        return false;
    });

    function previewDocument(fullUrlPath) {
        function previewReport(fullUrlPath) {
            $("#preview_report_container").empty();
            var loaderImg = $("<img></img>");
            $(loaderImg).attr("src", "images/ajax-loader-files.gif");
            $(loaderImg).attr("style", "margin-left:" + $("#preview_report_container").width() / 2 + "px;margin-top:" + $("#preview_report_container").height() / 2 + "px");
            $("#preview_report_container").append(loaderImg);
            if (fullUrlPath.indexOf(".png", 0) > -1 || fullUrlPath.indexOf(".jpeg", 0) > -1 || fullUrlPath.indexOf(".jpg", 0) > -1 || fullUrlPath.indexOf(".gif", 0) > -1) {
                $("#preview_report_container").empty();
                $("#preview_report_container").append("<img src='" + fullUrlPath + "'/>");
            } else if (fullUrlPath.indexOf(".pdf", 0) > -1) {
                var viewObj = $("<object></object>");
                $(viewObj).attr("data", fullUrlPath);
                $(viewObj).attr("width", $("#pn_previewReport").width() + "px");
                $(viewObj).attr("height", ($("#preview_report_container").height() - 7) + "px");
                $("#preview_report_container").empty();
                $("#preview_report_container").append(viewObj);
            } else if (fullUrlPath.indexOf(".doc", 0) > -1 || fullUrlPath.indexOf(".docx", 0) > -1 || fullUrlPath.indexOf(".txt", 0) > -1 || fullUrlPath.indexOf(".xls", 0) > -1 || fullUrlPath.indexOf(".xlsx", 0) > -1 || fullUrlPath.indexOf(".csv", 0) > -1) {
                var officeConverstionDataObj = { theFileUrl: fullUrlPath };
                $.ajax({
                    type: "post",
                    contentType: "application/json; charset=utf-8",
                    url: GroupMJsConfig.webServicePath + "/getOfficeHtml5View",
                    dataType: "json",
                    dataFilter: function (data) {
                        return data;
                    },
                    data: '{"postOfficeObj":' + JSON.stringify(officeConverstionDataObj) + '}',
                    success: function (data) {
                        if (data.officeConverstionDataObj.convertedHtml5FileUrl !== null) {
                            $("#preview_report_container").empty();
                            if (fullUrlPath.indexOf(".csv", 0) == -1) {
                                var iframeObj = $("<iframe></iframe>");
                                iframeObj.attr("src", data.officeConverstionDataObj.convertedHtml5FileUrl);
                                iframeObj.attr("width", $("#preview_report_container").width());
                                iframeObj.attr("height", $("#preview_report_container").height() - 10);
                                iframeObj.attr("style", "border:0px");
                                $("#preview_report_container").append(iframeObj);
                            } else {
                                $("#preview_report_container").append(data.officeConverstionDataObj.fileContext);
                            }
                        } else {
                            $("#preview_report_container").append("<div class='alert alert-block alert-danger' style='padding:5px;'>" + data.officeConverstionDataObj.error + "</div>");
                        }
                    },
                    error: function (jqXHR, textStatus, errorThrown) {
                        $("#preview_report_container").append("<div class='alert alert-block alert-danger' style='padding:5px;'>Client request error:" + errorThrown + "</div>");
                    }

                });
            } else {
                $("#preview_report_container").empty();
                $("#preview_report_container").append("<div class='alert alert-block alert-danger' style='padding:5px;'>This file type can not be previewed.</div>");
            }
        }
    }
});

$(window).load(function () {
    $("#myViewer").css("height", ($(document).height() - $("#viewer_controller").height()) + 'px');
});