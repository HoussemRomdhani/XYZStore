﻿(function () {
    var appUrl = GetUrlKeyValue("SPAppWebUrl");
    var hostUrl = GetUrlKeyValue("SPHostUrl");
    var webRepo = new XYZ.Repositories.WebRepository();

    jQuery(function () {
        var message = jQuery("#message");

        var call = webRepo.getProperties(appUrl)
        call.done(function (data, textStatus, jqXHR) {
            var currentVersion = data.d['CurrentVersion'];

            if (SP.ScriptUtility.isNullOrEmptyString(currentVersion) == false) {
                populateInterface();
            } else {
                var call1 = webRepo.getPermissions(appUrl);
                var call2 = webRepo.getPermissions(appUrl, hostUrl);
                var calls = jQuery.when(call1, call2);
                calls.done(function (appResponse, hostResponse) {
                    var appPerms = new SP.BasePermissions();
                    appPerms.initPropertiesFromJson(appResponse[0].d.EffectiveBasePermissions);
                    var hostPerms = new SP.BasePermissions();
                    hostPerms.initPropertiesFromJson(hostResponse[0].d.EffectiveBasePermissions);
                    var manageWeb = appPerms.has(SP.PermissionKind.manageWeb);
                    var manageLists = hostPerms.has(SP.PermissionKind.manageLists);

                    if ((manageWeb && manageLists) === false) {
                        message.text("A site owner needs to visit this site to process an update");
                    } else {
                        message.text("Provisioning content to App Web");

                        var prov = new XYZ.Provisioner(appUrl, hostUrl);
                        var call = prov.execute();
                        call.progress(function (msg) {
                            message.append("<br/>");
                            message.append(msg);
                        });
                        call.done(function () {
                            setTimeout(function () {
                                populateInterface();
                            }, 4000);
                        });
                        call.fail(failHandler);
                    }

                });
                call.fail(failHandler);
            }
        });
        call.fail(failHandler);
    });

    function populateInterface() {
        var prodRepo = new XYZ.Repositories.ProductRepository(appUrl, hostUrl);
        var call = prodRepo.getProductsByCategory("Beverages");
        call.done(function (data, textStatus, jqXHR) {
            var message = jQuery("#message");
            message.text("Products:");
            jQuery.each(data.d.results, function (index, value) {
                message.append("<br/>");
                message.append(value.Title);
            });
        });
        call.fail(failHandler);
    }


    function failHandler(errObj) {
        var response = "";
        if (errObj.get_message) {
            response = errObj.get_message();
        } else {
            try {
                var parsed = JSON.parse(errObj.responseText);
                response = parsed.error.message.value;
            } catch (e) {
                response = errObj.responseText;
            }
        }
        alert("Call failed. Error: " + response);
    }
})();