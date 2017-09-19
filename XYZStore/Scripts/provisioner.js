window.XYZ = window.XYZ || {};

XYZ.Provisioner = function (appUrl, hostUrl) {
    var dfd;

    function createProductsList() {
        if (!dfd) return;
        dfd.notify("Creating Products list");

        var context = new SP.ClientContext(appUrl);
        var web = context.get_web();

        var lci = new SP.ListCreationInformation();
        lci.set_title("Products");
        lci.set_templateType(SP.ListTemplateType.genericList);
        var list = web.get_lists().add(lci);

        context.executeQueryAsync(success, fail);

        function success() {
            dfd.notify("Products list created");
            dfd.resolve();
        }

        function fail(sender, args) {
            dfd.reject(args);
        }
    }

    function execute() {
        dfd = new jQuery.Deferred();

        createProductsList();

        return dfd.promise();
    }

    return {
        execute: execute
    }
}