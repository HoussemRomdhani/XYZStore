window.XYZ = window.XYZ || {};

XYZ.Provisioner = function (appUrl, hostUrl) {
    var dfd;
    var categoryListId;

    function createCategoriesList() {
        if (!dfd) return;
        dfd.notify("Creating Categories list");

        var context = new SP.ClientContext(appUrl);
        var web = XYZ.Repositories.getWeb(context, hostUrl);

        var lci = new SP.ListCreationInformation();
        lci.set_title("Categories");
        lci.set_templateType(SP.ListTemplateType.genericList);
        var list = web.get_lists().add(lci);

        list.get_fields().addFieldAsXml('<Field DisplayName="Description" Type="Text" Required="FALSE" Name="Description" />', true, SP.AddFieldOptions.defaultValue);

        context.load(list);
        context.executeQueryAsync(success, fail);

        function success() {
            dfd.notify("Categories list created");

            categoryListId = list.get_id();
            getCategoriesData();
        }

        function fail(sender, args) {
            dfd.reject(args);
        }
    }

    function getCategoriesData() {
        if (!dfd) return;
        dfd.notify("Requesting Categories data");

        var url = appUrl + "/Content/CategoriesData.txt";
        var call = jQuery.get(url);
        call.done(function (data, textStatus, jqXHR) {
            populateCategoriesList(data);
        });
        call.fail(function (jqXHR, textStatus, errorThrown) {
            dfd.reject(jqXHR);
        });
    }

    function populateCategoriesList(data) {
        if (!dfd) return;
        dfd.notify("Populating Categories list");

        var context = new SP.ClientContext(appUrl);
        var web = XYZ.Repositories.getWeb(context, hostUrl);
        var list = web.get_lists().getByTitle("Categories");

        var categories = JSON.parse(data);
        for (var i = 0; i < categories.length; i++) {
            var value = categories[i];
            var ici = new SP.ListItemCreationInformation();
            var item = list.addItem(ici);
            item.set_item("Title", value.Title);
            item.set_item("Description", value.Description);
            item.update();
        };

        context.executeQueryAsync(success, fail);

        function success() {
            dfd.notify("Categories list populated");
            createProductsList();
        }

        function fail(sender, args) {
            dfd.reject(args);
        }
    }



    function createProductsList() {
        if (!dfd) return;
        dfd.notify("Creating Products list");

        var context = new SP.ClientContext(appUrl);
        var web = XYZ.Repositories.getWeb(context, hostUrl);

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
            getProductsData();
        }

        function fail(sender, args) {
            dfd.reject(args);
        }
    }

    function getProductsData() {
        if (!dfd) return;
        dfd.notify("Requesting Products data");

        var url = appUrl + "/Content/ProductsData.txt";
        var call = jQuery.get(url);
        call.done(function (data, textStatus, jqXHR) {
            populateProductsList(data);
        });
        call.fail(function (jqXHR, textStatus, errorThrown) {
            dfd.reject(jqXHR);
        });
    }

    function populateProductsList(data) {
        if (!dfd) return;
        dfd.notify("Populating Products list");

        var context = new SP.ClientContext(appUrl);
        var web = XYZ.Repositories.getWeb(context, hostUrl);
        var list = web.get_lists().getByTitle("Products");

        var products = JSON.parse(data);
        var currentProduct = -1;
        updateNextSet();

        function updateNextSet() {
            var setIndex = 0;
            while (true) {
                setIndex += 1;
                currentProduct += 1;
                if (setIndex == 25 || currentProduct == products.length) {
                    context.executeQueryAsync(success, fail);
                    break;
                }

                var value = products[currentProduct];
                var ici = new SP.ListItemCreationInformation();
                var item = list.addItem(ici);
                item.set_item("Title", value.Title);
                item.set_item("QuantityPerUnit", value.QuantityPerUnit);
                item.set_item("UnitPrice", value.UnitPrice);
                item.set_item("UnitsInStock", value.UnitsInStock);
                item.set_item("UnitsOnOrder", value.UnitsOnOrder);
                item.set_item("ReorderLevel", value.ReorderLevel);
                item.set_item("Discontinued", value.Discontinued);

                var lfv = new SP.FieldLookupValue();
                lfv.set_lookupId(value.CategoryId);
                item.set_item("Category", lfv);
                item.update();
            };
        }

        function success() {
            dfd.notify(String.format("\t{0} of {1}", currentProduct, products.length));
            if (currentProduct == products.length) {
                updateCurrentVersion();
            } else {
                updateNextSet();
            }
        }

        function fail(sender, args) {
            dfd.reject(args);
        }

    }

    function updateCurrentVersion() {
        if (!dfd) return;
        dfd.notify("Updating current version number");

        var repo = new XYZ.Repositories.WebRepository();
        var call = repo.setProperty("CurrentVersion", "1.0.0.0", appUrl);
        call.done(function (data, textStatus, jqXHR) {
            dfd.notify("Update complete");
            dfd.resolve();
        });
        call.fail(function (jqXHR, textStatus, errorThrown) {
            dfd.reject(jqXHR);
        });
    }


    function execute() {
        dfd = new jQuery.Deferred();

        createCategoriesList();

        return dfd.promise();
    }

    return {
        execute: execute
    }
}