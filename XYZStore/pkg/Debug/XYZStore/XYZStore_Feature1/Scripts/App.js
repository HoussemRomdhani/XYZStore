(function () {
    "use strict";

    var appUrl = GetUrlKeyValue("SPAppWebUrl");
    var hostUrl = GetUrlKeyValue("SPHostUrl");

    jQuery(document).ready(function () {
        jQuery("#addWebPartFilesButton").click(addWebPartFiles);
    });

    function addWebPartFiles() {
        UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);

        var call1 = copyFile("WebPartContent", "Site Assets", "databinding.txt");
        var call2 = copyFile("WebPartContent", "Web Part Gallery", "databinding.dwp");
        var calls = jQuery.when(call1, call2);
        calls.done(function (response1, response2) {
            var message = jQuery("#message");
            message.text("Web Part files copied");
        });
        calls.fail(failHandler);
    }

    function getFile(sourceFolder, fileName) {
        var fileUrl = String.format("{0}/{1}/{2}",
            _spPageContextInfo.webServerRelativeUrl, sourceFolder, fileName);
        var call = jQuery.ajax({
            url: appUrl + "/_api/Web/GetFileByServerRelativeUrl('" + fileUrl + "')/$value",
            type: "GET",
            headers: {
                Accept: "text/plain"
            }
        });

        return call;
    }

    function uploadFile(fileContents, targetLibrary, fileName) {
        var url = String.format("{0}/_api/SP.AppContextSite(@target)" +
            "/Web/Lists/getByTitle('{1}')/RootFolder/Files/Add(url='{2}',overwrite=true)?@target='{3}'",
            appUrl, targetLibrary, fileName, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "POST",
            data: fileContents,
            processData: false,
            headers: {
                Accept: "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                "content-length": fileContents.length
            }
        });

        return call;
    }

    function copyFile(sourceFolder, targetLibrary, fileName) {
        var call = getFile(sourceFolder, fileName)
            .then(function (fileContents) { return uploadFile(fileContents, targetLibrary, fileName) });

        return call;
    }

    function failHandler(jqXHR, textStatus, errorThrown) {
        var response = "";
        try {
            var parsed = JSON.parse(jqXHR.responseText);
            response = parsed.error.message.value;
        } catch (e) {
            response = jqXHR.responseText;
        }
        alert("Call failed. Error: " + response);
    }



})();
