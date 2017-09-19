window.XYZ = window.XYZ || {};
window.XYZ.Repositories = window.XYZ.Repositories || {};

XYZ.Repositories.getWeb = function (context, hostUrl) {
    var web = null;

    if (hostUrl) {
        var hostContext = new SP.AppContextSite(context, hostUrl);
        web = hostContext.get_web();
    } else {
        web = context.get_web();
    }

    return web;
}

XYZ.Repositories.targetUrl = function (url, hostUrl) {
    if (hostUrl) {
        var api = "_api/";
        var index = url.indexOf(api);
        url = url.slice(0, index + api.length) +
            "SP.AppContextSite(@target)" +
            url.slice(index + api.length - 1);

        var connector = "?";
        if (url.indexOf("?") > -1 && url.indexOf("$") > -1) {
            connector = "&";
        }

        url = url + connector + "@target='" + hostUrl + "'";
    }

    return url;
}

XYZ.Repositories.WebRepository = function () {
    function getProperties(appUrl, hostUrl) {
        var url = appUrl + "/_api/Web/AllProperties";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    function setProperty(name, value, appUrl, hostUrl) {
        var dfd = new jQuery.Deferred();

        var context = SP.ClientContext.get_current();
        var web = XYZ.Repositories.getWeb(context, hostUrl);
        var props = web.get_allProperties();

        props.set_item(name, value);
        web.update();
        context.executeQueryAsync(success, fail);

        function success() {
            dfd.resolve();
        }

        function fail(sender, args) {
            dfd.reject(args);
        }

        return dfd.promise();
    }

    function getPermissions(appUrl, hostUrl) {
        var url = appUrl + "/_api/Web/effectiveBasePermissions";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    return {
        getProperties: getProperties,
        setProperty: setProperty,
        getPermissions: getPermissions
    }
}

XYZ.Repositories.CategoryRepository = function (appUrl) {
    var listUrl = "/_api/Web/Lists/getByTitle('Categories')";

    function getCategories() {
        var url = appUrl + listUrl + "/Items?$select=Id,Title";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    return {
        getCategories: getCategories
    }
}

XYZ.Repositories.ProductRepository = function (appUrl, hostUrl) {
    var listUrl = "/_api/Web/Lists/getByTitle('Products')";

    function getProducts(orderby, top) {
        if (!orderby) orderby = "Id";
        if (!top) top = 15;

        var url = appUrl + listUrl + "/Items?$select=*,Category/Title&$orderby=" + orderby + "&$top=" + top + "&$expand=Category/Title";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    function getProductsByCategory(category) {
        if (!category) category = "Beverages";

        var url = appUrl + listUrl + "/Items?$select=*,Category/Title&$filter=(Category/Title eq '" + category + "')&$expand=Category/Title";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    function getProduct(id) {
        if (!id) id = "1";

        var url = appUrl + listUrl + "/Items(" + id + ")?$select=*,Category/Title&$expand=Category/Title";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    function saveProduct(id, data, formDigest) {
        var url = appUrl + listUrl + "/Items(" + id + ")";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "POST",
            data: data,
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": formDigest,
                "IF-MATCH": "*",
                "X-Http-Method": "PATCH"
            }
        });

        return call;
    }

    function addProduct(data, formDigest) {
        var url = appUrl + listUrl + "/Items";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "POST",
            data: data,
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": formDigest
            }
        });

        return call;
    }

    function getPermissions() {
        var url = appUrl + listUrl + "/effectiveBasePermissions";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        return call;
    }

    function getNextProductId() {
        var dfd = new jQuery.Deferred();

        var url = appUrl + listUrl + "/Items?$top=1&$select=ProductID&$orderby=ProductID desc";
        url = XYZ.Repositories.targetUrl(url, hostUrl);

        var call = jQuery.ajax({
            url: url,
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        call.done(function (data, textStatus, jqXHR) {
            var productId = 1;

            if (data.d.results.length == 1) {
                productId = data.d.results[0].ProductID + 1;
            }

            dfd.resolve(productId);
        });
        call.fail(function (jqXHR, textStatus, errorThrown) {
            dfd.resolve(0);
        });

        return dfd.promise();
    }

    return {
        getProducts: getProducts,
        getProductsByCategory: getProductsByCategory,
        getProduct: getProduct,
        saveProduct: saveProduct,
        addProduct: addProduct,
        getPermissions: getPermissions,
        getNextProductId: getNextProductId
    }
}