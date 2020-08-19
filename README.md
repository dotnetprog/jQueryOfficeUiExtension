# jQueryOfficeUiExtension
a jQuery plugin that extents fabric js components (with this , you can build editable grid from the table component)


## DataSource

The Datasource object is a must to use for using the officeui extensions.

Currently odata is only supported as datasource type

here's some exemples
```javascript
var dt_src = new OfficeUi.dataSource({
              type: 'odata',
              queryOptions: {
                  $filter: 'InvoiceId eq ' + $("#Id").val(),
                  $expand: 'Product'
              },
              url: '/odata/InvoiceLine',
              async: true,
              schema: {
                  key: 'Id'
              },
          });
var products_dt_src = new OfficeUi.dataSource({
    type: 'odata',
    url: '/odata/Products',
    async: false

});
```

## editable table fabric js

here's an exemple on how to setup the plugin
```javascript
var dt_src = new OfficeUi.dataSource({
              type: 'odata',
              queryOptions: {
                  $filter: 'InvoiceId eq ' + $("#Id").val(),
                  $expand: 'Product'
              },
              url: '/odata/InvoiceLine',
              async: true,
              schema: {
                  key: 'Id'
              },
          });
var inv_grid = $('#lineGrid').OfficeEditableTable({
                datasource: dt_src,
                minheight: "300px",
                commandbar: [
                    {
                        type: 'create'

                    },
                    {
                        type: "delete"
                    },
                    {
                        type: 'save'
                    }

                ],
                columns: [
                    {
                        width: '400px',
                        label: 'Product',
                        field: 'Productid',
                        hidden: false,
                        template: function (dataRow) {

                            var pn = "(none)";
                            products_dt_src.fetch(function (products) {
                                pn = products.filter(function (p) { return p.Id === dataRow['Productid']; })[0].Name;

                            });

                            return pn;


                        },
                        editor: {
                            template: function (container, dataRow) {
                                var dropdown = $(container).OfficeUIDropdown({
                                    datasource: products_dt_src,
                                    labelkey: 'Name',
                                    valuekey: 'Id',
                                    selectedvalue: !!dataRow.Productid ? dataRow.Productid : null,
                                    onChange: function (e) {
                                        var productid = $(e.currentTarget).val();
                                        dataRow.set('Productid', productid);
                                    }
                                });

                                if (!dataRow.Productid) {
                                    var opt = dropdown.getValue();
                                    dataRow.set('Productid', opt.value);
                                }


                            }
                        }
                    },
                    {
                        label: 'InvoiceId',
                        hidden: true,
                        field: 'Invoiceid',
                        defaultvalue: $("#Id").val()
                    },
                    {
                        label: 'Price Per Unit',
                        field: 'Priceperunit',
                        type: 'decimal',
                        format: {
                            fixed: 2
                        },
                        hidden: false
                    },
                    {
                        label: 'Quantity',
                        field: 'Quantity',
                        type: 'decimal',
                        format: {
                            fixed: 2
                        },
                        hidden: false
                    },
                    {
                        label: 'Total Amount',

                        calculated: function (dataRow) {
                            if (!!dataRow.Priceperunit && !!dataRow.Quantity)
                                return dataRow.Priceperunit * dataRow.Quantity;

                            return 0;
                        },
                        triggerFields: ['Quantity', 'Priceperunit'],
                        type: 'decimal',
                        format: 'money',
                        hidden: false
                    }
                ],
                selectable: true,
                IsReadOnly: false,

            });
        });
```
## Dropdown Component fabric js

```javascript
var products_dt_src = new OfficeUi.dataSource({
    type: 'odata',
    url: '/odata/Products',
    async: false

});
 var dropdown = $('#container').OfficeUIDropdown({
                                    datasource: products_dt_src,
                                    labelkey: 'Name',
                                    valuekey: 'Id',
                                    selectedvalue: !!dataRow.Productid ? dataRow.Productid : null,
                                    onChange: function (e) {
                                        var productid = $(e.currentTarget).val();
                                        dataRow.set('Productid', productid);
                                    }
                                });
//how to get selected value of the dropdown
var opt = dropdown.getValue();
var opt_label = opt.label;
var opt_value = opt.value;
```

documentation on going....


# MAKE SURE YOU USE THE FABRIC JS FROM THE DEPENDENCIES DIRECTORY , SOME FIXES HAS BEEN ADDED TO THIS ONE.
