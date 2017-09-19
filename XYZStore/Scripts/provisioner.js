window.XYZ = window.XYZ || {};

XYZ.Provisioner = function (appUrl, hostUrl) {
    var dfd;

    function getCategoryListId() {
        if (!dfd) return;
        dfd.notify("Getting Category list id");

        var context = new SP.ClientContext(appUrl);
        var web = context.get_web();

        var list = web.get_lists().getByTitle("Categories");

        context.load(list, "Id");
        context.executeQueryAsync(success, fail);

        function success() {
            createProductsList(list.get_id());
        }

        function fail(sender, args) {
            dfd.reject(args);
        }
    }

    function createProductsList(categoryListId) {
        if (!dfd) return;
        dfd.notify("Creating Products list");

        var context = new SP.ClientContext(appUrl);
        var web = context.get_web();

        var lci = new SP.ListCreationInformation();
        lci.set_title("Products");
        lci.set_templateType(SP.ListTemplateType.genericList);
        var list = web.get_lists().add(lci);

        list.get_fields().addFieldAsXml('<Field DisplayName="Category" Type="Lookup" Required="FALSE" List="{' + categoryListId + '}" Name="Category" ShowField="Title" Version="1" />', true, SP.AddFieldOptions.defaultValue);
        list.get_fields().addFieldAsXml('<Field DisplayName="QuantityPerUnit" Type="Text" Required="FALSE" Name="QuantityPerUnit" />', true, SP.AddFieldOptions.defaultValue);
        list.get_fields().addFieldAsXml('<Field DisplayName="UnitPrice" Type="Currency" Required="FALSE" Name="UnitPrice" />', true, SP.AddFieldOptions.defaultValue);
        list.get_fields().addFieldAsXml('<Field DisplayName="UnitsInStock" Type="Integer" Required="FALSE" Name="UnitsInStock" />', true, SP.AddFieldOptions.defaultValue);
        list.get_fields().addFieldAsXml('<Field DisplayName="UnitsOnOrder" Type="Integer" Required="FALSE" Name="UnitsOnOrder" />', true, SP.AddFieldOptions.defaultValue);
        list.get_fields().addFieldAsXml('<Field DisplayName="ReorderLevel" Type="Integer" Required="FALSE" Name="ReorderLevel" />', true, SP.AddFieldOptions.defaultValue);
        list.get_fields().addFieldAsXml('<Field DisplayName="Discontinued" Type="Boolean" Required="FALSE" Name="Discontinued" />', true, SP.AddFieldOptions.defaultValue);


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

        getCategoryListId();

        return dfd.promise();
    }

    return {
        execute: execute
    }
}