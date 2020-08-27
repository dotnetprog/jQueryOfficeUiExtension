# jQueryOfficeUiExtension
a jQuery plugin that extents fabric js components (with this , you can build editable grid from the table component)

Features: 
- Column Sorting
- Column Resizing
- Column Templating
- Column Formatting
- Calculated Field/Column
- Commandbar customizable
- CRUD Operations with Save changes
- Inline Editing !
- Paging
- Scroll with fixed header

Coming soon:
- UI Column Filtering
- New datasource type "custom"
- Plugin's Api Documentation & how to use
## Demo

https://screencast-o-matic.com/watch/cYjObgmB4o

## DataSource

The Datasource object is a must to use for using the officeui extensions.

Currently odata is only supported as datasource type

### Api Reference
\* = required. <br>
OfficeUi.dataSource constructor 
- paramaters
  - Object with following properties
    - type*:string (only 'odata' is supported at the moment)
    - queryOptions:object containing all the odata queryOptions/queryStrings for the url
    - url*:string represent the endpoint to call to retrieve data
    - async:boolean if the call is async or sync
    - schema(*required only when used on table):object defines the object model that the datasource will handle, mainly used to determine the primary key of the record.
    - odata:object defines the odata setting , currently used to setup the count odata call for the footer if displayed.



### here's some exemples
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

### Api Reference
\* = required. <br>
$.fn.OfficeEditableTable 
- paramaters
  - Object with following properties
    - datasource*:OfficeUi.dataSource Object 
    - commandbar:array of object type command
    - columns:array of object type column
    - selctable:boolean row can be selected or not
    - IsReadOnly*:boolean disable the edit feature
    - paging:object defines the pagination of the gird/table
      - size:number number of records per page
      - displayCount: display the sum total rows. if size is defined, the index is shown as well.

$.fn.OfficeEditableTable command
- Object
  - type:string
    - Available types:
      - create (uses the datasource.create)
      - delete (uses the datasource.delete)
      - save (uses the datasource.saveChanges)
      - refresh
      - custom
  - onClick:function fired when button is clicked, works with custom
  - label:string label for button of type 'custom'
  - icon:string fabric ui icon to use (works only with custom)
  

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
### Paging
```javascript
 var inv_grid = null;
        $(document).ready(function () {

            var dt_src = new OfficeUi.dataSource({
                type: 'odata',
                odata: {
                    counturl:'/odata/Invoices('+$("#Id").val()+')/InvoiceLine/$count' //this config is used to display total records and for paging
                },
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
           

            inv_grid = $('#lineGrid').OfficeEditableTable({
                datasource: dt_src,
                commandbar: [
                   //define your commands here...
                ],
                columns: [
                    //define columns here...
                ],
                paging: {
                    displayCount: true,//shows the count in footer , if datasource type = odata, counturl on odata config of datasource must be set
                    size: 5 //determines how many records per page.
                },
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
