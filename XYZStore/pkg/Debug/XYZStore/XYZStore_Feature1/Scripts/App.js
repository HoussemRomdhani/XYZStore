(function () {
    "use strict";

    var appUrl = GetUrlKeyValue("SPAppWebUrl");
    var hostUrl = GetUrlKeyValue("SPHostUrl");

    jQuery(document).ready(function () {
        jQuery("#addWebPartFilesButton").click(addWebPartFiles);
        jQuery("#removeWebPartFilesButton").click(removeWebPartFiles);
    });

    function addWebPartFiles() {
        UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);

        var webPartFileName = "databinding.dwp";
        var call1 = copyFile("WebPartContent", "Site Assets", "databinding.txt");
        var call2 = getFile("WebPartContent", webPartFileName)
            .then(updateContentLink)
            .then(function (fileContents) { return uploadFile(fileContents, "Web Part Gallery", webPartFileName) })
            .then(updateWebPartGroup)
        var calls = jQuery.when(call1, call2);
        calls.done(function (response1, response2) {
            var message = jQuery("#message");
            message.text("Web Part files copied");
        });
        calls.fail(failHandler);
    }

    function removeWebPartFiles() {
        UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);

        var call1 = deleteFile("Site Assets", "databinding.txt");
        var call2 = deleteFile("Web Part Gallery", "databinding.dwp");
        var calls = jQuery.when(call1, call2);
        calls.done(function (response1, response2) {
            var message = jQuery("#message");
            message.text("Web Part files removed");
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
            "/Site/RootWeb/Lists/getByTitle('{1}')/RootFolder/Files/Add(url='{2}',overwrite=true)?@target='{3}'",
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

    function updateWebPartGroup(data) {
        var file = data.d;
        var url = String.format("{0}/_api/SP.AppContextSite(@target)" +
            "/Site/RootWeb/GetFileByServerRelativeUrl('{1}')/ListItemAllFields?" +
            "@target='{2}'",
            appUrl, file.ServerRelativeUrl, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "POST",
            data: JSON.stringify({
                "__metadata": { type: "SP.Data.OData__x005f_catalogs_x002f_wpItem" },
                Group: "App Script Part"
            }),
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                "IF-MATCH": "*",
                "X-Http-Method": "PATCH"
            }
        });

        return call;
    }

    function updateContentLink(fileContents) {
        var def = new jQuery.Deferred();

        var fileUrl = hostUrl + "/SiteAssets/databinding.txt";
        fileContents = fileContents.replace("{ContentLink}", fileUrl);
        def.resolve(fileContents);

        return def.promise();
    }

    function deleteFile(targetLibrary, fileName) {
        var url = String.format("{0}/_api/SP.AppContextSite(@target)" +
            "/Site/RootWeb/Lists/getByTitle('{1}')/RootFolder/Files('{2}')?@target='{3}'",
            appUrl, targetLibrary, fileName, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "POST",
            headers: {
                Accept: "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                "IF-MATCH": "*",
                "X-Http-Method": "DELETE"
            }
        });

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
